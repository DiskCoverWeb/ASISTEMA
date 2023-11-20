VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FPuntoVenta 
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   17880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   17880
   WindowState     =   1  'Minimized
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
      Height          =   5580
      Left            =   1365
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   8100
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "PuntoVenta.frx":0000
         DataSource      =   "AdoBenef"
         Height          =   5250
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   9260
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
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   5895
      Begin MSDataListLib.DataList DLGrupo 
         Bindings        =   "PuntoVenta.frx":0017
         DataSource      =   "AdoGrupo"
         Height          =   2940
         Left            =   105
         TabIndex        =   16
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
   Begin VB.TextBox TxtEmail 
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
      MaxLength       =   60
      TabIndex        =   5
      ToolTipText     =   "Escriba el Nombre/C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   735
      Width           =   6735
   End
   Begin VB.TextBox TxtDesc 
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
      Left            =   9975
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "PuntoVenta.frx":002E
      Top             =   1890
      Width           =   1065
   End
   Begin VB.TextBox TxtServicio 
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
      Left            =   8820
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "PuntoVenta.frx":0030
      Top             =   1890
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Facturar Mesa"
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
      Left            =   15435
      TabIndex        =   53
      Top             =   1575
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mesa"
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
      Left            =   14595
      Picture         =   "PuntoVenta.frx":0032
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1575
      Width           =   750
   End
   Begin VB.CommandButton CButtonSalir 
      Caption         =   "SALI&R"
      Height          =   855
      Left            =   11235
      Picture         =   "PuntoVenta.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7140
      Width           =   1275
   End
   Begin VB.CommandButton CButtonGrabar 
      Caption         =   "&GRABAR"
      Height          =   855
      Left            =   11235
      Picture         =   "PuntoVenta.frx":0C06
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6195
      Width           =   1275
   End
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
      Left            =   9030
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   45
      Text            =   "PuntoVenta.frx":14D0
      Top             =   7245
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
      Left            =   3465
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "PuntoVenta.frx":14D7
      Top             =   7245
      Width           =   2115
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
      TabIndex        =   22
      Text            =   "PuntoVenta.frx":14DE
      Top             =   1890
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
      Left            =   11130
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "PuntoVenta.frx":14E0
      Top             =   1890
      Width           =   1485
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "PuntoVenta.frx":14E5
      DataSource      =   "AdoBodega"
      Height          =   420
      Left            =   1470
      TabIndex        =   14
      Top             =   1050
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
      Left            =   11025
      TabIndex        =   31
      Text            =   "0"
      Top             =   525
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
      TabIndex        =   9
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   420
      Value           =   -1  'True
      Width           =   645
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "PuntoVenta.frx":14FD
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   20
      Top             =   1890
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
      Bindings        =   "PuntoVenta.frx":1517
      Height          =   3795
      Left            =   105
      TabIndex        =   48
      Top             =   2310
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   6694
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
      Top             =   3150
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
      Top             =   3465
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
      Top             =   3465
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
      Top             =   3150
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
      Top             =   3150
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
      Top             =   3465
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
      Top             =   3150
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
      Top             =   3150
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
   Begin VB.Label Label18 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EMAIL:"
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
      TabIndex        =   4
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Descuento"
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
      TabIndex        =   25
      Top             =   1575
      Width           =   1065
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
      Left            =   12705
      TabIndex        =   29
      Top             =   1890
      Width           =   1800
   End
   Begin VB.Label LabelServicio 
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
      Left            =   3465
      TabIndex        =   55
      Top             =   7770
      Width           =   2115
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Servicio"
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
      TabIndex        =   23
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total &Servicio"
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
      Left            =   105
      TabIndex        =   54
      Top             =   7770
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Total"
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
      Left            =   12705
      TabIndex        =   49
      Top             =   1575
      Width           =   1800
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
      Left            =   9030
      TabIndex        =   47
      Top             =   7770
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
      Left            =   9030
      TabIndex        =   43
      Top             =   6720
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
      Left            =   9030
      TabIndex        =   41
      Top             =   6195
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
      Left            =   3465
      TabIndex        =   39
      Top             =   8295
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
      Left            =   5670
      TabIndex        =   46
      Top             =   7770
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
      Left            =   5670
      TabIndex        =   44
      Top             =   7245
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
      Left            =   5670
      TabIndex        =   42
      Top             =   6720
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
      Left            =   5670
      TabIndex        =   40
      Top             =   6195
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
      Left            =   105
      TabIndex        =   38
      Top             =   8295
      Width           =   3375
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
      Left            =   105
      TabIndex        =   36
      Top             =   7245
      Width           =   3375
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
      TabIndex        =   18
      Top             =   1050
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
      TabIndex        =   17
      Top             =   1050
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
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   105
      Width           =   1380
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
      TabIndex        =   21
      Top             =   1575
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
      Left            =   11130
      TabIndex        =   27
      Top             =   1575
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
      TabIndex        =   19
      Top             =   1575
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
      TabIndex        =   13
      Top             =   1050
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
      Left            =   11025
      TabIndex        =   30
      Top             =   105
      Width           =   2115
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
      Left            =   3465
      TabIndex        =   35
      Top             =   6720
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
      Left            =   3465
      TabIndex        =   33
      Top             =   6195
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
      Left            =   105
      TabIndex        =   34
      Top             =   6720
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
      Left            =   105
      TabIndex        =   32
      Top             =   6195
      Width           =   3375
   End
End
Attribute VB_Name = "FPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQLFA As String
Dim SQLDecFA As String

Private Sub CButtonGrabar_Click()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  FechaTexto = MBFecha
  CalculosTotalesFactura AdoAsientoF
  If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then
     'FA.Cod_CxC = "CXC" & TipoFactura
     'Lineas_De_CxC FA
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
        SelectDataGrid DGAsientoF, AdoAsientoF, sSQLFA, SQLDecFA
        Ln_No = 1
       'Encerar_Facturas
'       FA.Cod_CxC = "CXC" & TipoFactura
 '      Lineas_De_CxC FA
        FA.Factura = Numero_Factura(FA)
        TextFacturaNo.Text = Format(FA.Factura, "0000000")
        DCArticulo.SetFocus
     End If
  Else
     MsgBox "Error: El Efectivo no alcanza para grabar"
  End If
End Sub

''Private Sub CButtonPCs_Click(Index As Integer)
''  If TotalPCs(Index) > 0 Then
''     With AdoAsientoF.Recordset
''      If .RecordCount Then
''         .MoveFirst
''         .Find ("CODIGO = '99.02." & Format(Index + 1, "000") & "' ")
''          If Not .EOF Then
''            .Delete
''            .Update
''          End If
''      End If
''     End With
''     SetAddNew AdoAsientoF
''     SetFields AdoAsientoF, "CODIGO", "99.02." & Format(Index + 1, "000")
''     SetFields AdoAsientoF, "CODIGO_L", CodigoL
''     SetFields AdoAsientoF, "PRODUCTO", "PC No." & Format(Index + 1, "00") & ", T[" & TiempoPCs(Index) & "]"
''     SetFields AdoAsientoF, "CANT", 1
''     SetFields AdoAsientoF, "PRECIO", TotalPCs(Index)
''     SetFields AdoAsientoF, "TOTAL", TotalPCs(Index)
''     SetFields AdoAsientoF, "Total_IVA", 0
''     SetFields AdoAsientoF, "RUTA", TiempoPCs(Index)
''     SetFields AdoAsientoF, "Item", NumEmpresa
''     SetFields AdoAsientoF, "CodigoU", CodigoUsuario
''     SetFields AdoAsientoF, "A_No", Ln_No
''     SetUpdate AdoAsientoF
''     Ln_No = Ln_No + 1
''  End If
''  CalculosTotalesFactura AdoAsientoF
''End Sub

''Private Sub CButtonPCs_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
''  Keys_Especiales Shift
''  If KeyCode = vbKeyEscape Then TxtEfectivo.SetFocus
''End Sub

Private Sub CButtonSalir_Click()
  Cyber_Cabinas = False
  Unload FPuntoVenta
End Sub

Private Sub Command1_Click()
   FPedidos.Show 1
End Sub

Private Sub Command2_Click()
''   Mensajes = "LLENAR PEDIDOS" & vbCrLf _
''            & "DESDE CERO"
''   Titulo = "FORMULARIO DE ELIMINACION"
''   If BoxMensaje = vbYes Then
''      Ln_No = 0
''      sSQL = "DELETE * " _
''           & "FROM Asiento_F " _
''           & "WHERE Item = '" & NumEmpresa & "' " _
''           & "AND CodigoU = '" & CodigoUsuario & "' "
''      ConectarAdoExecute sSQL
''   End If
   Cadena = DCCliente.Text & vbCrLf & vbCrLf _
          & "ORDEN No."
   Habitacion_No = UCase$(InputBox(Cadena, "FACTURACION DE PEDIDOS", ""))
   If Habitacion_No = "" Then Habitacion_No = Ninguno
   
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND HABIT = '" & Habitacion_No & "' "
   ConectarAdoExecute sSQL
   
   sSQL = "SELECT TP.Codigo_Sup,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0," _
        & "SUM(TP.Cantidad) As Cant,AVG(TP.PRECIO) As PVP,SUM(TP.Total) As VTotal," _
        & "SUM(TP.Total_IVA) As VTotal_IVA " _
        & "FROM Trans_Pedidos As TP,Catalogo_Productos As CP " _
        & "WHERE TP.Item = '" & NumEmpresa & "' " _
        & "AND TP.Periodo = '" & Periodo_Contable & "' " _
        & "AND TP.Orden_No = " & Val(Habitacion_No) & " " _
        & "AND TP.Codigo_Sup = CP.Codigo_Inv " _
        & "AND TP.Item = CP.Item " _
        & "AND TP.Periodo = CP.Periodo " _
        & "GROUP BY TP.Codigo_Sup,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0 " _
        & "ORDER BY TP.Codigo_Sup "
   SelectAdodc AdoAsientoF, sSQL
   With AdoAsientoF.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Codigo = .Fields("Codigo_Sup")
           Codigo1 = .Fields("Producto")
           If FA.TC = "NV" Then
              Cta = .Fields("Cta_Ventas_0")
           Else
              Cta = .Fields("Cta_Ventas")
           End If
           Precio = .Fields("PVP")
           Total = .Fields("VTotal")
           Total_IVA = .Fields("VTotal_IVA")
           Cantidad = .Fields("Cant")
           Insertar_Pedidos
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT TP.Codigo,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0," _
        & "SUM(TP.Cantidad) As Cant,AVG(TP.PRECIO) As PVP,SUM(TP.Total) As VTotal," _
        & "SUM(TP.Total_IVA) As VTotal_IVA " _
        & "FROM Trans_Pedidos As TP,Catalogo_Productos As CP " _
        & "WHERE TP.Item = '" & NumEmpresa & "' " _
        & "AND TP.Periodo = '" & Periodo_Contable & "' " _
        & "AND TP.Orden_No = " & Val(Habitacion_No) & " " _
        & "AND CP.Agrupacion = " & Val(adFalse) & " " _
        & "AND TP.Codigo = CP.Codigo_Inv " _
        & "AND TP.Item = CP.Item " _
        & "AND TP.Periodo = CP.Periodo " _
        & "GROUP BY TP.Codigo,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0 " _
        & "ORDER BY TP.Codigo "
   SelectAdodc AdoAsientoF, sSQL
   With AdoAsientoF.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Codigo = .Fields("Codigo")
           Codigo1 = .Fields("Producto")
           If FA.TC = "NV" Then
              Cta = .Fields("Cta_Ventas_0")
           Else
              Cta = .Fields("Cta_Ventas")
           End If
           Precio = .Fields("PVP")
           Total = .Fields("VTotal")
           Total_IVA = .Fields("VTotal_IVA")
           Cantidad = .Fields("Cant")
           Insertar_Pedidos
          .MoveNext
        Loop
    End If
   End With
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQLFA, SQLDecFA
  CalculosTotalesFactura AdoAsientoF
  DCArticulo.SetFocus
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then
     CodigoBenef = Ninguno
     TextBenef = Ninguno
     CodigoCliente = Ninguno
     NombreCliente = Ninguno
     Grupo_No = Ninguno
     TipoDoc = Ninguno
     DCCliente.Text = Ninguno
     FrmBenef.Visible = False
     FA.CodigoC = CodigoCliente
     FA.Cliente = NombreCliente
     TextBenef.SetFocus
  End If
End Sub

Private Sub DCCliente_LostFocus()
  If DCCliente.Text <> Ninguno Then
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
          FA.Cliente = NombreCliente
          TextCotiza.SetFocus
       Else
          NombreCliente = DCCliente.Text
          FrmBenef.Visible = False
          FPuntoVenta.Visible = False
          Nuevo = True
          MsgBox "Cliente No existe"
          FClientesFlash.Show 1
          FPuntoVenta.Visible = True
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
  End If
End Sub

Private Sub DGAsientoF_KeyDown(KeyCode As Integer, Shift As Integer)
  'If KeyCode = vbKeyDelete Then CalculosTotalesFactura AdoAsientoF
End Sub

Private Sub DGAsientoF_LostFocus()
  CalculosTotalesFactura AdoAsientoF
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

''Private Sub MBTiempo_KeyDown(KeyCode As Integer, Shift As Integer)
''  PresionoEnter KeyCode
''End Sub

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
          FA.EmailC = .Fields("Email")
          Grupo_No = .Fields("Grupo")
          TipoDoc = .Fields("TD")
          FA.CodigoC = CodigoCliente
          FA.Cliente = NombreCliente
          TxtEmail.Text = FA.EmailC
          TextCotiza.SetFocus
       Else
         .MoveFirst
         .Find ("CI_RUC = '" & TextBenef & "' ")
          If Not .EOF Then
             CodigoBenef = .Fields("Codigo")
             TextBenef.Text = .Fields("Cliente")
             CodigoCliente = .Fields("Codigo")
             NombreCliente = .Fields("Cliente")
             FA.EmailC = .Fields("Email")
             Grupo_No = .Fields("Grupo")
             TipoDoc = .Fields("TD")
             FA.CodigoC = CodigoCliente
             FA.Cliente = NombreCliente
             TxtEmail.Text = FA.EmailC
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
  sSQL = "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,Email " _
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
  Total_Servicio = 0: Total_IVA = 0
  NumLn = 0
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQLFA, SQLDecFA
  With DtaProd.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total_IVA = Total_IVA + .Fields("Total_IVA")
          Total_Desc = Total_Desc + .Fields("Total_Desc")
          Total_Servicio = Total_Servicio + .Fields("SERVICIO")
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
  LabelServicio.Caption = Format(Total_Servicio, "#,##0.00")
  Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
  FA.SubTotal = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc, 2)
  FA.Total_IVA = Total_IVA
  FA.Total_MN = Total_Factura
  FA.Con_IVA = Total_Con_IVA
  FA.Sin_IVA = Total_Sin_IVA
  FA.Servicio = Total_Servicio
  FA.Descuento = Total_Desc
  LabelSubTotal.Caption = Format(Total_Sin_IVA, "#,##0.00")
  LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0.00")
  LabelIVA.Caption = Format(Total_IVA, "#,##0.00")
  TxtDescuento = Format(Total_Desc, "#,##0.00")
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
  Total_FacturaME = 0
  If Val(TextCotiza) > 0 Then
     TotalDolar = Val(CCur(TextCotiza))
     If OpcDiv.value Then
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
  DatInv.Patron_Busqueda = DCArticulo
  If Leer_Codigo_Inv(DCArticulo, FechaSistema, Cod_Bodega) Then
     DatosArticulos
  Else
     ListaStock.Show 1
     DatosArticulos
  End If
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
  Mensajes = "¿Realmente desea eliminar el campo " & vbCrLf & "(" _
           & AdoAsientoF.Recordset.Fields("CODIGO") & ") " _
           & AdoAsientoF.Recordset.Fields("PRODUCTO") & "   TOTAL -> " _
           & AdoAsientoF.Recordset.Fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = vbYes Then
     Cancel = False
     DCArticulo.SetFocus
  Else
     Cancel = True
  End If
End Sub

Private Sub Form_Activate()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  FA.TC = TipoFactura
  FA.Fecha = MBFecha
  Ln_No = 1
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
  FA.Cod_CxC = "CXC" & TipoFactura
  If TipoFactura = "PV" Then
     FPuntoVenta.Caption = "INGRESAR TICKET"
     Label1.Caption = " TICKET No."
     Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  ElseIf TipoFactura = "NV" Then
     FPuntoVenta.Caption = "INGRESAR NOTA DE VENTA"
     Label1.Caption = " NOTA DE VENTA No."
     Label3.Caption = "I.V.A. 0.00%"
  ElseIf TipoFactura = "OP" Then
     FPuntoVenta.Caption = "INGRESAR ORDER DE PRODUCCION"
     Label1.Caption = " ORDEN No."
     Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  Else
     FPuntoVenta.Caption = "INGRESAR FACTURA"
     Label1.Caption = " FACTURA No."
     Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  End If
  FPuntoVenta.Caption = FPuntoVenta.Caption & " (" & FA.Cod_CxC & ")"
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  
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
       & "AND Fact = '" & TipoFactura & "' " _
       & "AND Codigo = '" & FA.Cod_CxC & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "ORDER BY Codigo "
  SelectAdodc AdoLinea, sSQL
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  With AdoLinea.Recordset
    If .RecordCount > 0 Then
        Cta_Cobrar = .Fields("CxC")
        CodigoL = .Fields("Codigo")
        FA.Cod_CxC = .Fields("Codigo")
    Else
        MsgBox "Falta Organizar la CxC en Puntos de Venta." & vbCrLf _
             & "Salga de este proceso y llame al su técnico" & vbCrLf _
             & "o al Contador de su Organizacion."
    End If
  End With
  Lineas_De_CxC FA
  FA.Factura = Numero_Factura(FA)
  TextFacturaNo = FA.Factura
  TextFacturaNo.Text = Format(FA.Factura, "0000000")
  
  VerSiExisteCta Cta_CajaG
  VerSiExisteCta Cta_CajaGE
  VerSiExisteCta Cta_CajaBA
  sSQLFA = "SELECT CODIGO,CANT,PRODUCTO,PRECIO,Total_Desc,Total_IVA,TOTAL,SERVICIO,VALOR_TOTAL,Orden_No,Mes,Cod_Ejec," _
         & "Porc_C,REP,FECHA,CODIGO_L,HABIT,RUTA,TICKET,Cta,Cta_SubMod,Item,CodigoU,CodBod,CodMar,TONELAJE,CORTE,A_No," _
         & "Codigo_Cliente,Numero,Serie,Autorizacion,Codigo_B,PRECIO2,Total_Desc2 " _
         & "FROM Asiento_F " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
  SQLDecFA = "PRECIO 4|,CORTE 5|."
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQLFA, SQLDecFA
  
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
  sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'I' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDBList DLGrupo, AdoGrupo, sSQL, "Producto"
  
  sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "AND Mid$(Codigo_Inv,1,2) <> '99' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDBCombo DCArticulo, AdoArticulo, sSQL, "Producto"
  RatonNormal
  FPuntoVenta.WindowState = 2
  If AdoArticulo.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Productos de Venta"
     Unload FPuntoVenta
  Else
     Listar_Clientes
  End If
End Sub

Private Sub Form_Deactivate()
  FPuntoVenta.WindowState = 1
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
  Encerar_Facturas
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  If TipoFactura = "PV" Then
     FPuntoVenta.Caption = "INGRESAR TICKET"
     Label1.Caption = " TICKET No."
     Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  ElseIf TipoFactura = "NV" Then
     FPuntoVenta.Caption = "INGRESAR NOTA DE VENTA"
     Label1.Caption = " NOTA DE VENTA No."
     Label3.Caption = "I.V.A. 0.00%"
  Else
     FPuntoVenta.Caption = "INGRESAR FACTURA"
     Label1.Caption = " FACTURA No."
     Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  End If
  FPuntoVenta.Caption = FPuntoVenta.Caption & " (" & FA.Cod_CxC & ")"
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  FechaTexto1 = MBFecha.Text
End Sub

Private Sub TextCant_Change()
  Real1 = Val(TextCant) * Val(TextVUnit)
  LabelVTotal.Caption = Format(Real1, "#,##0.0000")
End Sub

Private Sub TextCant_GotFocus()
  If Val(TextVUnit) <= 0 Then TextVUnit = Format(Precio, "#,##0.0000")
  MarcarTexto TextCant
  TxtServicio = "0"
  TxtDesc = "0"
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
   If Len(DatInv.Cta_Inventario) <= 1 Then Grabar_PV = True
   'MsgBox Cant_Item_PV
   If Grabar_PV Then
      TextoValido TextCant, True
      LabelVTotal.Caption = Format(Real1, "#,##0.00")
      Real2 = 0: Real3 = 0
      Real1 = CCur(Val(TextCant)) * CCur(Val(TextVUnit))
      Real1 = Redondear(Real1, 2)
      With AdoAsientoF.Recordset
        If Real1 > 0 Then
           Real2 = Redondear(Real1 * (Val(TxtDesc) / 100), 2)
           If BanIVA Then Real3 = (Real1 - Real2) * Porc_IVA Else Real3 = 0
           If EsNotaVenta Then Real3 = 0
           Real5 = Redondear((Real1 - Real2) * (Val(TxtServicio) / 100), 2)
           SetAddNew AdoAsientoF
           SetFields AdoAsientoF, "CODIGO", Codigos
           SetFields AdoAsientoF, "CODIGO_L", CodigoL
           SetFields AdoAsientoF, "PRODUCTO", Mid$(Producto, 1, 150)
           SetFields AdoAsientoF, "CANT", CDbl(TextCant)
           SetFields AdoAsientoF, "PRECIO", CDbl(TextVUnit)
           SetFields AdoAsientoF, "TOTAL", Real1
           SetFields AdoAsientoF, "Total_IVA", Real3
           SetFields AdoAsientoF, "Total_Desc", Real2
           SetFields AdoAsientoF, "SERVICIO", Real5
           SetFields AdoAsientoF, "VALOR_TOTAL", Real1 + Real3 - Real2 + Real5
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
  TextoValido TxtEfectivo, True, , 2
 'Seteamos los encabezados para las facturas
  CalculosTotalesFactura AdoAsientoF
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     If Total_Factura > 200 And CodigoCliente = "9999999999" Then
        MsgBox "No se puede grabar Facturas a Consumidor Final con este monto"
     Else
        RatonReloj
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
       SelectAdodc AdoAsientoF, sSQL
        With AdoAsientoF.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                sSQL = "DELETE * " _
                     & "FROM Trans_Pedidos " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Orden_No = " & Val(.Fields("HABIT")) & " "
                ConectarAdoExecute sSQL
               .MoveNext
             Loop
         End If
        End With
   '     FA.Cod_CxC = "CXC" & TipoFactura
   '     Lineas_De_CxC FA
        FA.Factura = Numero_Factura(FA)
        FA.Fecha_C = FA.Fecha
        FA.Fecha_V = FA.Fecha
        FA.T = Cancelado
        TextoFormaPago = PagoCred
       'Grabamos el numero de factura
        Grabar_Factura FA, True
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
           TA.CodigoC = FA.CodigoC
           TA.Serie = FA.Serie
           FA.Efectivo = CCur(TxtEfectivo)
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
        SRI_Crear_Clave_Acceso_Facturas FA, False, True
        'MsgBox Grafico_PV
        If FA.TC = "OP" Then
           FA.Desde = FA.Factura
           FA.Hasta = FA.Factura
           FA.Tipo_PRN = "OP"
           Imprimir_Facturas_CxC FPuntoVenta, FA
        Else
           If Grafico_PV Then
              Imprimir_Punto_Venta_Grafico FA
           Else
              Imprimir_Punto_Venta FA
           End If
        End If
        SelectDataGrid DGAsientoF, AdoAsientoF, sSQLFA, SQLDecFA
        NombreCliente = "CONSUMIDOR FINAL"
        TextBenef = "CONSUMIDOR FINAL"
        TxtEfectivo = "0.00"
        CodigoBenef = "9999999999"
        CodigoCliente = "9999999999"
        FA.CodigoC = CodigoCliente
        FA.Cliente = NombreCliente
        TextBenef.SetFocus
     End If
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub DatosArticulos()
  If DatInv.Codigo_Inv <> Ninguno Then
     LabelStock.Caption = DatInv.Stock
     Codigos = DatInv.Codigo_Inv
     Producto = DatInv.Producto
     Cta_Ventas = DatInv.Cta_Ventas
     Precio = Redondear(DatInv.PVP, 2)
     BanIVA = DatInv.IVA
    'MsgBox DatInv.IVA
     TextVUnit = Format(Precio, "#,##0.0000")
     If EsNotaVenta Then BanIVA = False
     DCArticulo.Text = Producto  '& " -> " & Codigos
     'MsgBox DCArticulo.Text
  End If
End Sub

'''Public Sub DatosArticulos()
'''  With AdoArticulo.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''      .Find ("Codigo_Inv = '" & DatInv.Codigo_Inv & "' ")
'''       If Not .EOF Then
'''          DatInv.Producto = .Fields("Producto")
'''          LabelStock.Caption = DatInv.Stock
'''          Codigos = DatInv.Codigo_Inv
'''          Producto = DatInv.Producto
'''          Cta_Ventas = DatInv.Cta_Ventas
'''          Precio = Redondear(DatInv.PVP, 2)
'''          BanIVA = DatInv.IVA
'''         'MsgBox DatInv.IVA
'''          TextVUnit = Format(Precio, "#,##0.0000")
'''          If EsNotaVenta Then BanIVA = False
'''          DCArticulo.Text = Producto  '& " -> " & Codigos
'''       End If
'''   End If
'''  End With
'''End Sub

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
''''          Codigo = Mid$(Codigo, Len(Codigo) - 1, 2)
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

Private Sub TxtDesc_GotFocus()
  MarcarTexto TxtDesc
End Sub

Private Sub TxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc_LostFocus()
  TextoValido TxtDesc
End Sub

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
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(TxtEfectivo) - Total_FacturaME, "#,##0.00")
     FA.Total_MN = Total_FacturaME
  Else
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(TxtEfectivo) - Total_Factura, "#,##0.00")
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

Public Sub Insertar_Pedidos()
   If Len(Cta) > 1 And Len(Codigo) > 1 Then
      SetAdoAddNew "Asiento_F"
      SetAdoFields "CODIGO", Codigo
      SetAdoFields "CODIGO_L", CodigoL
      SetAdoFields "PRODUCTO", Codigo1
      SetAdoFields "CANT", Cantidad
      SetAdoFields "HABIT", Habitacion_No
      SetAdoFields "Orden_No", Val(Habitacion_No)
      SetAdoFields "PRECIO", Precio
      SetAdoFields "TOTAL", Total
      SetAdoFields "Total_IVA", Total_IVA
      SetAdoFields "Cta", Cta
      SetAdoFields "Item", NumEmpresa
      SetAdoFields "CodigoU", CodigoUsuario
      SetAdoFields "Cta_SubMod", FA.SubCta
      SetAdoFields "RUTA", TextComEjec
      SetAdoFields "CodBod", Cod_Bodega
      SetAdoFields "CodMar", Cod_Marca
      SetAdoFields "A_No", CByte(Ln_No)
''      If Val(TextComEjec.Text) > 0 Then
''         SetAdoFields "Cod_Ejec", CodigoEjec
''         SetAdoFields "Porc_C", Redondear(Val(TextComEjec) / 100, 4)
''      End If
      SetAdoUpdate
      Ln_No = Ln_No + 1
   End If
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  TextoValido TxtEmail
  If Len(TxtEmail) <= 1 Then TxtEmail = Lista_De_Correos(4).Correo_Electronico
  Actualiza_Email TxtEmail, CodigoCliente
  FA.EmailC = TxtEmail
  FA.EmailR = TxtEmail
End Sub

Private Sub TxtServicio_GotFocus()
  MarcarTexto TxtServicio
End Sub

Private Sub TxtServicio_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtServicio_LostFocus()
   TextoValido TxtServicio
End Sub
