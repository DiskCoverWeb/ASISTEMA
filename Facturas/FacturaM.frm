VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FacturasMult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   7065
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FacturaM.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   67
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
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
   Begin MSDataListLib.DataCombo DCEjecutivo 
      Bindings        =   "FacturaM.frx":001B
      DataSource      =   "AdoEjecutivo"
      Height          =   315
      Left            =   2280
      TabIndex        =   66
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Ejecutivo"
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FacturaM.frx":0036
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   1560
      TabIndex        =   65
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cliente"
   End
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "FacturaM.frx":004F
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   2400
      TabIndex        =   64
      Top             =   525
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Cliente"
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FacturaM.frx":0066
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   120
      TabIndex        =   62
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483646
      Text            =   "Producto"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   240
      Top             =   4680
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
   Begin MSAdodcLib.Adodc AdoBuscarCli 
      Height          =   330
      Left            =   240
      Top             =   4440
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
      Caption         =   "BuscarCli"
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   240
      Top             =   4920
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
      Caption         =   "Productos"
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
   Begin MSAdodcLib.Adodc AdoEjecutivo 
      Height          =   330
      Left            =   240
      Top             =   5160
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
      Caption         =   "Ejecutivo"
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
      Left            =   240
      Top             =   5400
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
   Begin MSAdodcLib.Adodc AdoTrans 
      Height          =   330
      Left            =   2280
      Top             =   4680
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
   Begin MSAdodcLib.Adodc AdoFact1 
      Height          =   330
      Left            =   2280
      Top             =   4920
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
      Caption         =   "Fact1"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   2280
      Top             =   5160
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
      Caption         =   "Cliente"
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
      Left            =   2280
      Top             =   5400
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   4320
      Top             =   4680
      Visible         =   0   'False
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoCodigos 
      Height          =   330
      Left            =   4320
      Top             =   4920
      Visible         =   0   'False
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
      Caption         =   "Codigos"
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
   Begin MSAdodcLib.Adodc AdoEjec 
      Height          =   330
      Left            =   4320
      Top             =   5160
      Visible         =   0   'False
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
      Caption         =   "Ejec"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4320
      Top             =   5400
      Visible         =   0   'False
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
      Caption         =   "Aux"
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
   Begin VB.TextBox TextGrupo 
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
      Left            =   6825
      MaxLength       =   15
      TabIndex        =   5
      Top             =   105
      Width           =   855
   End
   Begin VB.CheckBox CheqN 
      Caption         =   "Nueva"
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
      Top             =   2310
      Visible         =   0   'False
      Width           =   960
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "FacturaM.frx":0080
      Top             =   6615
      Width           =   1380
   End
   Begin VB.TextBox TextSector 
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   27
      Top             =   2310
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox TextArea 
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
      Left            =   10080
      MaxLength       =   5
      TabIndex        =   29
      Top             =   2310
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame FrameEjec 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   2730
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   8100
      Begin MSDataListLib.DataCombo DCEjec 
         Bindings        =   "FacturaM.frx":0085
         DataSource      =   "AdoEjec"
         Height          =   315
         Left            =   80
         TabIndex        =   63
         Top             =   450
         Width           =   5840
         _ExtentX        =   10319
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483646
         Text            =   "Ejecutivo"
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
      Begin VB.CommandButton Command5 
         BackColor       =   &H00808080&
         Caption         =   "Co&ntinuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6930
         TabIndex        =   40
         Top             =   210
         Width           =   1065
      End
      Begin VB.TextBox TextComEjec 
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
         Left            =   5985
         MaxLength       =   5
         TabIndex        =   39
         Text            =   "00.00"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Com. %"
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
         Left            =   5985
         TabIndex        =   38
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   " ELIJA EL EJECUTIVO DE VENTA"
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
         Left            =   105
         TabIndex        =   37
         Top             =   210
         Width           =   5790
      End
   End
   Begin VB.TextBox TextComision 
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
      Left            =   10080
      MaxLength       =   5
      TabIndex        =   20
      Text            =   "0"
      Top             =   1890
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox TextContrato 
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
      Left            =   5775
      MaxLength       =   8
      TabIndex        =   25
      Top             =   2310
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxHasta 
      Height          =   330
      Left            =   3255
      TabIndex        =   23
      Top             =   2310
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox MBoxDesde 
      Height          =   330
      Left            =   1995
      TabIndex        =   22
      Top             =   2310
      Visible         =   0   'False
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
   Begin VB.TextBox TextProducto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      MaxLength       =   120
      TabIndex        =   35
      Text            =   "x"
      Top             =   3360
      Width           =   10725
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
      Left            =   7770
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "FacturaM.frx":009B
      Top             =   3045
      Width           =   1485
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
      Left            =   6825
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "FacturaM.frx":00A0
      Top             =   3045
      Width           =   960
   End
   Begin VB.TextBox TextObs 
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
      Left            =   1575
      MaxLength       =   200
      TabIndex        =   17
      Top             =   1470
      Width           =   9255
   End
   Begin VB.TextBox TextNota 
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
      Left            =   1575
      MaxLength       =   200
      TabIndex        =   15
      Top             =   1155
      Width           =   9255
   End
   Begin VB.TextBox TextFacturaNo 
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
      TabIndex        =   7
      Text            =   "0"
      Top             =   105
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1470
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
   Begin VB.CheckBox CheqEjec 
      Caption         =   "Ejecutivo de Venta"
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
      Top             =   1890
      Width           =   2115
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
      Height          =   645
      Left            =   1050
      TabIndex        =   55
      Top             =   6300
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar &Factura"
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
      TabIndex        =   54
      Top             =   6300
      Width           =   855
   End
   Begin VB.TextBox TextRep 
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
      TabIndex        =   45
      Text            =   "FacturaM.frx":00A2
      Top             =   5775
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Factura en Modena Extranjera"
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
      Left            =   7770
      TabIndex        =   9
      Top             =   525
      Width           =   3060
   End
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   4725
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
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Grupo:"
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
      Left            =   6090
      TabIndex        =   4
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Vencimiento:"
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
      Left            =   2835
      TabIndex        =   2
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9135
      TabIndex        =   51
      Top             =   6615
      Width           =   1695
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Facturado"
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
      TabIndex        =   48
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label LabelIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7770
      TabIndex        =   50
      Top             =   6615
      Width           =   1380
   End
   Begin VB.Label LabelServ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6405
      TabIndex        =   60
      Top             =   6615
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7770
      TabIndex        =   52
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio 10%"
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
      Left            =   6405
      TabIndex        =   61
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total &Desc."
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
      TabIndex        =   56
      Top             =   6300
      Width           =   1380
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3570
      TabIndex        =   53
      Top             =   6615
      Width           =   1485
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sector"
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
      Left            =   6825
      TabIndex        =   26
      Top             =   2310
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Area"
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
      TabIndex        =   28
      Top             =   2310
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision %"
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
      Left            =   8925
      TabIndex        =   19
      Top             =   1890
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contrato No."
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
      Left            =   4515
      TabIndex        =   24
      Top             =   2310
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo"
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
      Left            =   1155
      TabIndex        =   21
      Top             =   2310
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LabelVTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   9240
      TabIndex        =   43
      Top             =   3045
      Width           =   1590
   End
   Begin VB.Label LabelStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   5880
      TabIndex        =   59
      Top             =   3045
      Width           =   960
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " OBSERVACION"
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
      TabIndex        =   16
      Top             =   1470
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOTA"
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
      Top             =   1155
      Width           =   1485
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
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
      TabIndex        =   10
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7980
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Pedido:"
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
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label LabelTelefono 
      BackColor       =   &H80000009&
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
      Left            =   9555
      TabIndex        =   13
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telefono:"
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
      Left            =   8505
      TabIndex        =   12
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUENTA POR COBRAR:"
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
      Top             =   525
      Width           =   2325
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
      Left            =   9240
      TabIndex        =   34
      Top             =   2730
      Width           =   1590
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
      Left            =   7770
      TabIndex        =   33
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rep."
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
      TabIndex        =   32
      Top             =   5775
      Visible         =   0   'False
      Width           =   750
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
      Left            =   6825
      TabIndex        =   31
      Top             =   2730
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock"
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
      Left            =   5880
      TabIndex        =   58
      Top             =   2730
      Width           =   960
   End
   Begin VB.Label LabelSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2100
      TabIndex        =   49
      Top             =   6615
      Width           =   1485
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total con IVA"
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
      Left            =   3570
      TabIndex        =   47
      Top             =   6300
      Width           =   1485
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total sin IVA"
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
      Left            =   2100
      TabIndex        =   46
      Top             =   6300
      Width           =   1485
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   7140
      TabIndex        =   11
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   30
      Top             =   2730
      Width           =   5790
   End
End
Attribute VB_Name = "FacturasMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqEjec_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CheqN_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxDesde_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxHasta_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextArea_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextComision_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextContrato_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextGrupo_GotFocus()
  MarcarTexto TextGrupo
End Sub

Private Sub TextGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextGrupo_LostFocus()
  TextoValido TextGrupo
  sSQL = "SELECT Codigo & '   ' & Cliente As NomCliente "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE T = 'C' "
  sSQL = sSQL & "AND FactM = True "
  If TextGrupo.Text <> Ninguno Then sSQL = sSQL & "AND Grupo = '" & TextGrupo.Text & "' "
  sSQL = sSQL & "ORDER BY Direccion,Cliente "
  SelectDBCombo DCCliente, AdoCliente, sSQL, "NomCliente", False
End Sub

Private Sub CheqEjec_Click()
 If CheqEjec.Value = 1 Then
    DCEjecutivo.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    Label33.Visible = True
    Label34.Visible = True
    TextArea.Visible = True
    TextSector.Visible = True
    TextComision.Visible = True
    TextContrato.Visible = True
    MBoxDesde.Visible = True
    MBoxHasta.Visible = True
    CheqN.Visible = True
    CheqN.Value = 1
 Else
    DCEjecutivo.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    Label13.Visible = False
    Label33.Visible = False
    Label34.Visible = False
    TextArea.Visible = False
    TextSector.Visible = False
    TextComision.Visible = False
    TextContrato.Visible = False
    MBoxDesde.Visible = False
    MBoxHasta.Visible = False
    CheqN.Visible = False
    CheqN.Value = 0
 End If
End Sub

Private Sub Command1_Click()
  Unload FacturasMult
End Sub

Private Sub Command2_Click()
   Mensajes = "Esta Seguro que desea grabar Facturas"
   Titulo = "Formulario de Grabacion"
   If BoxMensaje = 6 Then
      If Check1.Value = 1 Then Moneda_US = True Else Moneda_US = False
      CalculosTotales FacturasMult, AdoProductos
      TextoFormaPago = PagoCont
      ProcGrabar
   End If
   NumComp = ReadSetDataNum("Facturas", True, False)
   TextFacturaNo.Text = NumComp
End Sub

Private Sub Command5_Click()
  FrameEjec.Visible = False
  TextCant.SetFocus
End Sub

Private Sub DCEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCEjec_LostFocus()
 CodigoEjec = SinEspaciosIzq(DCEjec.Text)
 sSQL = "SELECT * FROM Clientes "
 sSQL = sSQL & "WHERE Codigo = '" & CodigoEjec & "' "
 SelectData AdoFact1, sSQL
 If AdoFact1.Recordset.RecordCount > 0 Then TextComEjec.Text = Format(AdoFact1.Recordset.Fields("Porc_C"), "#,##0.00")
End Sub

Private Sub DCEjecutivo_LostFocus()
 CodigoEjec = SinEspaciosIzq(DCEjecutivo.Text)
 sSQL = "SELECT * FROM Clientes "
 sSQL = sSQL & "WHERE Codigo = '" & CodigoEjec & "' "
 SelectData AdoFact1, sSQL
 If AdoFact1.Recordset.RecordCount > 0 Then TextComision.Text = Format(AdoFact1.Recordset.Fields("Porc_C"), "#,##0.00")
 
End Sub

Private Sub DCLinea_LostFocus()
 Si_No = False
 CodigoL = SinEspaciosIzq(DCLinea.Text)
 sSQL = "SELECT * FROM Linea_Producto "
 sSQL = sSQL & "WHERE Codigo = '" & CodigoL & "' "
 SelectData AdoFact1, sSQL
 With AdoFact1.Recordset
  If .RecordCount > 0 Then
      If .Fields("Cta_CxC") <> Ninguno Then Cta_Cobrar = .Fields("Cta_CxC")
      Si_No = .Fields("Todos")
  End If
 End With
 If Si_No Then
    sSQL = "SELECT Codigo & ' -> '& Articulo As Nom_Art " _
         & "FROM Articulo " _
         & "ORDER BY Articulo "
 Else
    sSQL = "SELECT Codigo & ' -> '& Articulo As Nom_Art " _
         & "FROM Articulo " _
         & "WHERE CodigoL = '" & CodigoL & "' " _
         & "ORDER BY Articulo "
 End If
 SelectDBCombo DCArticulo, AdoArticulo, sSQL, "Nom_Art", False
End Sub

Private Sub DCArticulo_GotFocus()
   LabelStock.Caption = ""
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
     Case vbKeyEscape
          Empleados = False
          Command2.SetFocus
     Case vbKeyReturn
          TextCant.SetFocus
   End Select
End Sub

Private Sub DCArticulo_LostFocus()
   Producto = ""
   Codigos = SinEspaciosIzq(DCArticulo.Text)
   sSQL = "SELECT * FROM Articulo "
   sSQL = sSQL & "WHERE Codigo = '" & Codigos & "' "
   SelectData AdoFact1, sSQL, False
   With AdoFact1.Recordset
   If .RecordCount > 0 Then
      Producto = .Fields("Articulo")
      TextVUnit.Text = Format(.Fields("PVP"), "#,##0.00")
      LabelStock.Caption = .Fields("Stock")
      If Porc_IVA > 0 Then BanIVA = .Fields("IVA") Else BanIVA = False
      Cadena = ""
      Mensajes = "Escriba el nombre del Empleado ?"
      If Empleados Then
         Cadena = InputBox(Mensajes, , "")
         Cadena = " (" & Cadena & ")"
      End If
   Else
      Mensajes = "Este Producto no existe," & Chr(13)
      Mensajes = Mensajes & "Repita la operacion."
      MsgBox Mensajes
      DCArticulo.SetFocus
   End If
   End With
End Sub

Private Sub DCCliente_GotFocus()
   MarcarTexto DCCliente
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
   Empleados = False
   CodigoCli = SinEspaciosIzq(DCCliente.Text)
   LabelCodigo.Caption = ""
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & " WHERE Codigo = '" & CodigoCli & "' "
   SelectData AdoAux, sSQL, False
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        LabelTelefono.Caption = .Fields("Telefono")
        LabelCodigo.Caption = .Fields("Codigo")
        CodigoCli = LabelCodigo.Caption
        DireccionCli = .Fields("Direccion")
    Else
        LabelCodigo.Caption = "Codigo"
    End If
   End With
End Sub

Private Sub DGDetalle_AfterDelete()
  CalculosTotales FacturasMult, AdoProductos
End Sub

Private Sub DGDetalle_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & Chr(13) & "(" _
           & AdoProductos.Recordset.Fields("CODIGO") & ") " _
           & AdoProductos.Recordset.Fields("PRODUCTO") & "   TOTAL -> " _
           & AdoProductos.Recordset.Fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Activate()
   
   Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
   Label36.Caption = "Serv. " & Format(Porc_Serv * 100, "#0.00") & "%"
   TextCant.Text = "0"
   TextRep.Text = "0"
   TextVUnit.Text = "0"
   LabelVTotal.Caption = "0"
   Modificar = False
   Bandera = True
   Mifecha = BuscarFecha(FechaSistema)
   sSQL = "SELECT Codigo & '  ' & Linea As LineaP "
   sSQL = sSQL & "FROM Linea_Producto "
   sSQL = sSQL & "ORDER BY Linea "
   SelectDBCombo DCLinea, AdoLinea, sSQL, "LineaP", False
   CodigoL = SinEspaciosIzq(DCLinea.Text)
   sSQL = "SELECT Codigo & ' ->'& Articulo As Nom_Art "
   sSQL = sSQL & "FROM Articulo "
   sSQL = sSQL & "WHERE CodigoL = '" & CodigoL & "' "
   sSQL = sSQL & "ORDER BY Articulo "
   SelectDBCombo DCArticulo, AdoArticulo, sSQL, "Nom_Art", False
   
   sSQL = "SELECT Codigo & '   ' & Cliente As NomCliente "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = 'C' "
   sSQL = sSQL & "AND FactM = True "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "NomCliente", False
   
   sSQL = "SELECT Codigo & '  ' & Cliente As NCliente "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = 'R' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DCEjecutivo, AdoEjecutivo, sSQL, "NCliente", False
   
   sSQL = "SELECT Codigo & '  ' & Cliente As NCliente "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = 'E' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBCombo DCEjec, AdoEjec, sSQL, "NCliente", False
   
   sSQL = "SELECT * FROM Detalles_" & CodigoUsuario & " "
   SelectData AdoProductos, sSQL
   sSQL = "DELETE * FROM Detalles_" & CodigoUsuario & " "
   DeleteData AdoProductos, sSQL
   SelectDataGrid DGDetalle, AdoProductos, "Detalles_" & CodigoUsuario & " "
   NumComp = ReadSetDataNum("Facturas", True, False)
   TextFacturaNo.Text = NumComp
   RatonNormal
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FacturasMult
   'Abriendo bases relacionadas
   ConectarAdodc AdoAux
   ConectarAdodc AdoEjec
   ConectarAdodc AdoLinea
   ConectarAdodc AdoCliente
   ConectarAdodc AdoEjecutivo
   ConectarAdodc AdoArticulo
   ConectarAdodc AdoProductos
   ConectarAdodc AdoBuscarCli
   ConectarAdodc AdoFact1
   ConectarAdodc AdoFactura
   ConectarAdodc AdoDetalle
   ConectarAdodc AdoTrans
   ConectarAdodc AdoCodigos
End Sub

Private Sub MBoxDesde_LostFocus()
  FechaValida MBoxDesde, False
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
   FechaTexto1 = MBoxFecha.Text
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV, False
End Sub

Private Sub MBoxHasta_LostFocus()
  FechaValida MBoxHasta, False
End Sub

Private Sub TextCant_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  If TextCant.Text = "" Then TextCant.Text = "0"
  Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
  LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextComEjec_GotFocus()
  MarcarTexto TextComEjec
End Sub

Private Sub TextComEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextComision_GotFocus()
  MarcarTexto TextComision
End Sub

Private Sub TextComision_LostFocus()
  TextoValido TextComision, True
End Sub

Private Sub TextContrato_GotFocus()
  MarcarTexto TextContrato
End Sub

Private Sub TextContrato_LostFocus()
  TextoValido TextContrato, False, True
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
  TextoValido TextDesc, True
  CalculosTotales FacturasMult, AdoProductos
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_LostFocus()
 If TextFacturaNo.Text = "" Then TextFacturaNo.Text = "0"
 Factura_No = Val(TextFacturaNo.Text)
 sSQL = "SELECT * FROM Facturas WHERE Factura = " & Factura_No & " "
 SelectData AdoFactura, sSQL, False
 If AdoFactura.Recordset.RecordCount > 0 Then
    MsgBox "Warning:" & Chr(13) & "Ya existe la Factura No. " & Format(Factura_No, "000000") & "."
    DCLinea.SetFocus
 End If
End Sub

Private Sub TextNota_GotFocus()
   TextNota.Text = ""
End Sub

Private Sub TextNota_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNota_LostFocus()
  TextoValido TextNota
End Sub

Private Sub TextObs_GotFocus()
  TextObs.Text = ""
End Sub

Private Sub TextObs_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextObs_LostFocus()
  TextoValido TextObs
End Sub

Private Sub TextProducto_GotFocus()
   TextProducto.Text = ""
End Sub

Private Sub TextProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextProducto_LostFocus()
   TextoValido TextProducto
   If CheqEjec.Value = 1 Then
      FrameEjec.Visible = True
      DCEjec.SetFocus
   End If
End Sub

Private Sub TextRep_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextRep_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextRep_LostFocus()
   If TextRep.Text = "" Then TextRep.Text = "0"
   TextRep.Text = Format(TextRep.Text, "#,##0.00")
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextSector_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_Change()
   'Real1 = CDbl(TextCant.Text) * CDbl(TextVUnit.Text)
   LabelVTotal.Caption = Format(Real1, "#,##0.00")
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
   If AdoProductos.Recordset.RecordCount <= 30 Then
     'Producto = DCArticulo.Text
      If TextProducto.Text <> Ninguno Then Producto = Producto & ", " & TextProducto.Text
      Real1 = CDbl(TextCant.Text) * CDbl(TextVUnit.Text)
      'Real2 = CDbl(TextRep.Text) * CDbl(TextVUnit.Text)
      LabelVTotal.Caption = Format(Real1, "#,##0.00")
      AdoProductos.Recordset.AddNew
      AdoProductos.Recordset.Fields("CODIGO") = Codigos
      AdoProductos.Recordset.Fields("CODIGO_L") = CodigoL
      AdoProductos.Recordset.Fields("PRODUCTO") = Mid(Producto, 1, 151)
      AdoProductos.Recordset.Fields("REP") = 0
      AdoProductos.Recordset.Fields("CANT") = CDbl(TextCant.Text)
      AdoProductos.Recordset.Fields("PRECIO") = CDbl(TextVUnit.Text)
      AdoProductos.Recordset.Fields("TOTAL") = Real1
      AdoProductos.Recordset.Fields("Total_Desc") = Real2
      If BanIVA Then Real3 = (Real1 - Real2) * Porc_IVA Else Real3 = 0
      AdoProductos.Recordset.Fields("Total_IVA") = Real3
      AdoProductos.Recordset.Fields("Cod_Ejec") = Ninguno
      AdoProductos.Recordset.Fields("Porc_C") = 0
      If CheqEjec.Value = 1 Then
         If CSng(TextComEjec.Text) > 0 Then
            AdoProductos.Recordset.Fields("Cod_Ejec") = SinEspaciosIzq(DCEjec.Text)
            AdoProductos.Recordset.Fields("Porc_C") = CSng(TextComEjec.Text)
         End If
      End If
      If ((Val(TextCant.Text) > 0) And (Codigos <> "")) Then AdoProductos.Recordset.Update
      CalculosTotales FacturasMult, AdoProductos
      TextVUnit.Text = ""
      DCArticulo.SetFocus
   Else
      MsgBox "Ya no se puede ingresar mas datos."
      Command1.SetFocus
   End If
End Sub

Public Sub ProcGrabar()
  'Seteamos los encabezados para las facturas
  If AdoProductos.Recordset.RecordCount > 0 Then
     RatonReloj
     TextoValido TextNota
     TextoValido TextObs
     TextoValido TextSector
     TextoValido TextArea, , True
     FechaValida MBoxDesde, False
     FechaValida MBoxHasta, False
     FechaValida MBoxFecha, False
     FechaValida MBoxFechaV, False
     TextoValido TextContrato, True
     If Check1.Value = 1 Then Moneda_US = True Else Moneda_US = False
     Asiento = ReadSetDataNum("Asiento", True, True)
     IngresoCaja = ReadSetDataNum("Ingreso Caja", True, True)
     SelectData AdoFactura, "Facturas", False
     SelectData AdoDetalle, "Detalle_Factura", False
    'SelectData AdoTrans, "Transacciones ", False
     FechaTexto = MBoxFecha.Text
     Total_FacturaME = 0
     CalculosTotales FacturasMult, AdoProductos
     If Moneda_US Then
        Total_Factura = Round((Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio) * Dolar)
        Total_FacturaME = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
     Else
        Total_Factura = Round(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio)
        Total_FacturaME = 0
     End If
     Saldo = Total_Factura
     Saldo_ME = Total_FacturaME
     DiarioCaja = 0
    'Procesamos las multiples facturas
     NumComp = ReadSetDataNum("Facturas", True, False)
     Factura_Desde = NumComp
     If AdoCliente.Recordset.RecordCount > 0 Then
        AdoCliente.Recordset.MoveFirst
        Do While Not AdoCliente.Recordset.EOF
           CodigoCli = SinEspaciosIzq(AdoCliente.Recordset.Fields("NomCliente"))
           Factura_No = ReadSetDataNum("Facturas", True, True)
           sSQL = "DELETE * FROM Detalle_Factura "
           sSQL = sSQL & "WHERE Factura_No = " & Factura_No & " "
           DeleteData AdoTrans, sSQL
           sSQL = "DELETE * FROM Facturas "
           sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
           DeleteData AdoTrans, sSQL
           sSQL = "DELETE * FROM Diario_Caja "
           sSQL = sSQL & "WHERE Factura = " & Factura_No & " "
           DeleteData AdoTrans, sSQL
           SelectData AdoTrans, "Diario_Caja ", False
           With AdoTrans.Recordset
               SetAddNew AdoTrans
               .Fields("Cotizacion") = 0
                If Moneda_US Then .Fields("Cotizacion") = Dolar
               SetFields AdoTrans, "T", Normal
               SetFields AdoTrans, "TP", ventas
               SetFields AdoTrans, " Fecha", FechaTexto
               SetFields AdoTrans, "Diario_No", DiarioCaja
               SetFields AdoTrans, "Caja_No", IngresoCaja
               SetFields AdoTrans, "Factura", Factura_No
               SetFields AdoTrans, "Monto_ME", Total_FacturaME
               SetFields AdoTrans, "Monto_MN", Total_Factura
               SetFields AdoTrans, "Caja_ME", 0
               SetFields AdoTrans, "Caja_MN", 0
               SetFields AdoTrans, "Caja_Vaucher", 0
               SetFields AdoTrans, "Abonos_ME", 0
               SetFields AdoTrans, "Abonos_MN", 0
               SetFields AdoTrans, "Saldo_MN", 0
               SetFields AdoTrans, "Saldo_ME", 0
               SetFields AdoTrans, "Codigo_C", CodigoCli
               SetFields AdoTrans, "CtaxCob", Cta_Cobrar
               SetFields AdoTrans, "CtaxVent", Cta_Ventas
               SetFields AdoTrans, "Saldo_ME", Saldo_ME
               SetFields AdoTrans, "Saldo_MN", Saldo
               SetUpdate AdoTrans
           End With
           If Saldo < 0 Then Saldo = 0
           If Saldo > 0 Then
              TextoFormaPago = PagoCred
              T = Pendiente
           Else
              TextoFormaPago = PagoCont
              T = Cancelado
           End If
           With AdoFactura.Recordset
               SetAddNew AdoFactura
               SetFields AdoFactura, "T", Pendiente
               SetFields AdoFactura, "ME", Moneda_US
               SetFields AdoFactura, "Factura", Factura_No
               SetFields AdoFactura, "Fecha", FechaTexto
               SetFields AdoFactura, "Fecha_C", FechaTexto
               SetFields AdoFactura, "Fecha_V", MBoxFechaV.Text
               SetFields AdoFactura, "Codigo_C", CodigoCli
               SetFields AdoFactura, "Vendedor", NombreUsuario
               SetFields AdoFactura, "Pedido_No", 0  'TextPedidos.Text
               SetFields AdoFactura, "Bultos_No", 0  'TextBultos.Text
               SetFields AdoFactura, "Gavetas_No", 0 'TextTransportador.Text
               SetFields AdoFactura, "Forma_Pago", TextoFormaPago
               SetFields AdoFactura, "Sin_IVA", Total_Sin_IVA
               SetFields AdoFactura, "Con_IVA", Total_Con_IVA
               SetFields AdoFactura, "SubTotal", Total_Sin_IVA + Total_Con_IVA
               SetFields AdoFactura, "Descuento", Total_Desc
               SetFields AdoFactura, "IVA", Total_IVA
               SetFields AdoFactura, "Servicio", Total_Servicio
               SetFields AdoFactura, "Total_MN", 0
               SetFields AdoFactura, "Total_ME", 0
               SetFields AdoFactura, "Comision", 0
                If Moneda_US Then
                  SetFields AdoFactura, "Total_ME", Total_FacturaME
                   Total = Total_FacturaME
                Else
                  SetFields AdoFactura, "Total_MN", Total_Factura
                   Total = Total_Factura
                End If
               SetFields AdoFactura, "Saldo_MN", Total_Factura
               SetFields AdoFactura, "Saldo_ME", Total_FacturaME
                If CheqEjec.Value = 1 Then
                  SetFields AdoFactura, "Cod_Ejec", SinEspaciosIzq(DCEjecutivo.Text)
                  SetFields AdoFactura, "Porc_C", CSng(TextComision.Text)
                Else
                  SetFields AdoFactura, "Cod_Ejec", Ninguno
                  SetFields AdoFactura, "Porc_C", 0
                End If
               SetFields AdoFactura, "Cotizacion", Dolar
               SetFields AdoFactura, "Observacion", TextObs.Text
               SetFields AdoFactura, "Nota", TextNota.Text
               SetFields AdoFactura, "Cta_CxC", Cta_Cobrar
               SetFields AdoFactura, "Cta_Venta", Cta_Ventas
               SetUpdate AdoFactura
           End With
           AdoProductos.Recordset.MoveFirst
           Do While Not AdoProductos.Recordset.EOF
              With AdoDetalle.Recordset
                  SetAddNew AdoDetalle
                  SetFields AdoDetalle, "T", Pendiente
                  SetFields AdoDetalle, "Factura_No", Factura_No
                  SetFields AdoDetalle, "Codigo_C", CodigoCli
                  SetFields AdoDetalle, "Fecha", FechaTexto
                  SetFields AdoDetalle, "Codigo", AdoProductos.Recordset.Fields("CODIGO")
                  SetFields AdoDetalle, "Cantidad", AdoProductos.Recordset.Fields("CANT")
                  SetFields AdoDetalle, "CodigoL", AdoProductos.Recordset.Fields("CODIGO_L")
                  SetFields AdoDetalle, "Reposicion", AdoProductos.Recordset.Fields("REP")
                  SetFields AdoDetalle, "Precio", AdoProductos.Recordset.Fields("PRECIO")
                  SetFields AdoDetalle, "Total", AdoProductos.Recordset.Fields("TOTAL")
                  SetFields AdoDetalle, "Total_Desc", AdoProductos.Recordset.Fields("Total_Desc")
                  SetFields AdoDetalle, "Total_IVA", AdoProductos.Recordset.Fields("Total_IVA")
                  SetFields AdoDetalle, "Producto", AdoProductos.Recordset.Fields("PRODUCTO")
                  SetFields AdoDetalle, "Cod_Ejec", AdoProductos.Recordset.Fields("Cod_Ejec")
                  SetFields AdoDetalle, "Porc_C", AdoProductos.Recordset.Fields("Porc_C")
                  SetUpdate AdoDetalle
              End With
              AdoProductos.Recordset.MoveNext
           Loop
           Factura_Hasta = Factura_No
           FacturasMult.Caption = "Codigo Cliente: " & CodigoCli & ", Factura Actual No. " & Factura_Hasta
           AdoCliente.Recordset.MoveNext
       Loop
       RatonNormal
     Else
        RatonNormal
        MsgBox "No existe Clientes para Facturar"
     End If
    'Grabamos el Numero de Factura
     sSQL = "DELETE * FROM Detalles_" & CodigoUsuario & " "
     DeleteData AdoProductos, sSQL
     sSQL = "SELECT * FROM Detalles_" & CodigoUsuario & " "
     SelectDataGrid DGDetalle, AdoProductos, sSQL
     RatonNormal
     ImprimirFacturasCxC FacturasMult, AdoFactura, AdoDetalle, Factura_Desde, Factura_Hasta
     FacturasMult.Caption = "FACTURACION: Ingreso de Facturas"
    'Cerrando las bases involucradas
     CloseData AdoFactura
     CloseData AdoDetalle
  Else
     MsgBox "No se puede grabar la Factura," & Chr(13) & "falta datos."
  End If
End Sub

