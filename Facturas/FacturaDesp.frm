VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturaDespensa 
   BackColor       =   &H00C0FFC0&
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12735
   Icon            =   "FacturaDesp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   12735
   WindowState     =   1  'Minimized
   Begin VB.CommandButton CmdCIBenef1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   10815
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   630
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
         Size            =   9.75
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
   Begin VB.TextBox TextVUnit 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   9450
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "FacturaDesp.frx":424A
      Top             =   1470
      Width           =   1380
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
      Left            =   8295
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FacturaDesp.frx":424F
      Top             =   1470
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "FacturaDesp.frx":4251
      DataSource      =   "AdoBodega"
      Height          =   360
      Left            =   3780
      TabIndex        =   3
      Top             =   105
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   49152
      ForeColor       =   16777215
      Text            =   "DC"
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
   Begin VB.TextBox TextFacturaNo 
      Enabled         =   0   'False
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
      Left            =   11340
      TabIndex        =   5
      Text            =   "000000000"
      Top             =   105
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FacturaDesp.frx":4269
      DataSource      =   "AdoArticulo"
      Height          =   360
      Left            =   105
      TabIndex        =   11
      Top             =   1470
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   10185
      Picture         =   "FacturaDesp.frx":4283
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5985
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11445
      Picture         =   "FacturaDesp.frx":458D
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5985
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "FacturaDesp.frx":4E57
      Height          =   3060
      Left            =   105
      TabIndex        =   20
      Top             =   1995
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   5398
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8454016
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
      Left            =   525
      Top             =   2205
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
      Left            =   525
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
      Left            =   525
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
      Left            =   525
      Top             =   2205
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   525
      Top             =   4620
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
      Left            =   525
      Top             =   4935
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   2625
      Top             =   2205
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   2625
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FacturaDesp.frx":4E71
      DataSource      =   "AdoCliente"
      Height          =   360
      Left            =   1365
      TabIndex        =   7
      ToolTipText     =   "<Ctrl+R>: Buscar por CI/RUC, <F12>: LLamar a la Historia Clinica, <Ctrl+F>: Listar Ordenes de Trabajo  del Cliente"
      Top             =   630
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " (P0) C.I. / R.U.C."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8925
      TabIndex        =   8
      Top             =   630
      Width           =   2010
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
      Left            =   10920
      TabIndex        =   19
      Top             =   1470
      Width           =   1695
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10920
      TabIndex        =   18
      Top             =   1155
      Width           =   1695
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio Unitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   9450
      TabIndex        =   16
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8295
      TabIndex        =   14
      Top             =   1155
      Width           =   1065
   End
   Begin VB.Label LabelStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7035
      TabIndex        =   13
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7035
      TabIndex        =   12
      Top             =   1155
      Width           =   1170
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " AFILIADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label LabelStockArt 
      BackColor       =   &H0000C000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   1155
      Width           =   6840
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LINEA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7245
      TabIndex        =   4
      Top             =   105
      Width           =   4110
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1065
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   7350
      TabIndex        =   22
      Top             =   6405
      Width           =   2745
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Despachado"
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
      Left            =   7350
      TabIndex        =   21
      Top             =   5985
      Width           =   2745
   End
End
Attribute VB_Name = "FacturaDespensa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
    Buscar_Cliente DCCliente.Text
End Sub

Private Sub DCCliente_LostFocus()
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       If IsNumeric(DCCliente.Text) Then .Find ("CI_RUC = '" & DCCliente.Text & "' ") Else .Find ("Cliente = '" & DCCliente.Text & "' ")
       If Not .EOF Then
          CodigoCliente = .fields("Codigo")
          Label3.Caption = " (P" & .fields("Cont_Salidas") & ") C.I. / R.U.C."
          TBeneficiario = Leer_Datos_Cliente_SP(CodigoCliente)
          FA.CodigoC = TBeneficiario.Codigo
          FA.Cliente = TBeneficiario.Cliente
          FA.TD = TBeneficiario.TD
          FA.CI_RUC = TBeneficiario.CI_RUC
          FA.TelefonoC = TBeneficiario.Telefono1
          FA.DireccionC = TBeneficiario.Direccion
          FA.EmailC = TBeneficiario.Email1
          FA.EmailC2 = TBeneficiario.Email2
          FA.EmailR = TBeneficiario.EmailR
          
          CodigoCliente = FA.CodigoC
          NombreCliente = FA.Cliente
          CmdCIBenef1.Caption = FA.CI_RUC
          DireccionCli = TBeneficiario.Direccion
          DireccionGuia = TBeneficiario.Direccion
          'SiguienteControl
          'If Mod_Fact Then TextFacturaNo.SetFocus Else TextObs.SetFocus
       Else
          MsgBox "Cliente no Asignado"
          DCCliente.SetFocus
       End If
   Else
       Nuevo = True
       NombreCliente = DCCliente.Text
       MsgBox "Cliente no Asignado"
       DCCliente.SetFocus
   End If
  End With
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Calculos_Totales_Factura(DtaProd As Adodc)
Dim NumLn As Byte
  Total_ME = 0
  Si_No = False
  Total_Factura = 0: Total_Desc = 0
  Total_Con_IVA = 0: Total_Sin_IVA = 0
  Total_IVA = 0
  NumLn = 0
  With DtaProd.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total_IVA = Total_IVA + .fields("Total_IVA")
          Total_Desc = Total_Desc + .fields("Total_Desc")
          If .fields("Total_IVA") > 0 Then
             Total_Con_IVA = Total_Con_IVA + .fields("TOTAL")
          Else
             Total_Sin_IVA = Total_Sin_IVA + .fields("TOTAL")
          End If
          NumLn = NumLn + 1
         .MoveNext
       Loop
   End If
  End With
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
  Total_FacturaME = 0
  TextCant.Text = ""
  LabelVTotal.Caption = ""
End Sub

Private Sub Command1_Click()
  Unload FacturaDespensa
End Sub

Private Sub Command3_Click()
    FechaValida MBFecha
    FechaTexto = MBFecha
    FA.Fecha = MBFecha
    FA.Fecha_V = MBFecha
    Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
             & "La Despensa No. " & TextFacturaNo.Text
    Titulo = "Formulario de Grabacion"
    If BoxMensaje = vbYes Then
        Moneda_US = False
        TextoFormaPago = PagoCont
        ProcGrabarDespensa
        
        Bandera = True
        Label1.Caption = "DESPENSA No. " & FA.Autorizacion & "-" & FA.Serie & "-"
        FA.Nuevo_Doc = True
        FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
        TextFacturaNo = Format(FA.Factura, "000000000")
        Ln_No = 1
       
        sSQL = "SELECT CP.Producto, CP.Codigo_Inv, CP.Codigo_Barra, SUM(TK.Entrada-TK.Salida) " _
             & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
             & "WHERE CP.Item = '" & NumEmpresa & "' " _
             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
             & "AND TK.CodBodega = '" & Cod_Bodega & "' " _
             & "AND CP.TC = 'P' " _
             & "AND CP.Item = TK.Item " _
             & "AND CP.Periodo = TK.Periodo " _
             & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
             & "GROUP BY TK.CodBodega, CP.Producto, CP.Codigo_Inv, CP.Codigo_Barra " _
             & "HAVING SUM(TK.Entrada - TK.Salida) > 0 " _
             & "ORDER BY CP.Producto, CP.Codigo_Inv "
        SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
       
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        SQLDec = "PRECIO 4|CORTE 5|."
        Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
        Buscar_Cliente Ninguno
        DCCliente.SetFocus
    End If
End Sub

Private Sub DCArticulo_GotFocus()
  Calculos_Totales_Factura AdoAsientoF
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
   Select Case KeyCode
     Case vbKeyEscape
          Calculos_Totales_Factura AdoAsientoF
          Command3.SetFocus
     Case vbKeyReturn
          TextCant.SetFocus
   End Select
End Sub

Private Sub DCArticulo_LostFocus()
  Codigos = Ninguno
  If Leer_Codigo_Inv(DCArticulo.Text, FechaSistema, Cod_Bodega) Then Datos_Articulos_Despensa
  If Val(LabelStock.Caption) <= 0 Then
     MsgBox "No se puede despachar este producto, no tiene existencia"
     DCArticulo.SetFocus
  End If
End Sub

Private Sub DGAsientoF_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & vbCrLf & "(" _
           & AdoAsientoF.Recordset.fields("CODIGO") & ") " _
           & AdoAsientoF.Recordset.fields("PRODUCTO") & "   TOTAL -> " _
           & AdoAsientoF.Recordset.fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Activate()
  MBFecha = FechaSistema
  Mifecha = BuscarFecha(FechaSistema)
  FechaValida MBFecha
  FA.TC = TipoFactura
  FA.Fecha = MBFecha
  
  FacturaDespensa.Caption = "(" & FA.TC & ") "
  sSQL = "SELECT Codigo, Usuario, Clave, Nombre_Completo, TODOS, Ok, CodBod, EmailUsuario, Serie_FA " _
       & "FROM Accesos " _
       & "WHERE Codigo = '" & CodigoUsuario & "' "
  Select_Adodc AdoGrupo, sSQL
 'MsgBox sSQL & vbCrLf & .RecordCount
  If AdoGrupo.Recordset.RecordCount > 0 Then
     Cod_Bodega = AdoGrupo.Recordset.fields("CodBod")
     FA.Serie = AdoGrupo.Recordset.fields("Serie_FA")
     FA.Cod_CxC = "DE999" & MidStrg(FA.Serie, 4, 3)
     If Cod_Bodega <> Ninguno Then
         FA.Autorizacion = "1111111111"
         Lineas_De_CxC FA
         If FA.Cta_CxP <> Ninguno Then
            FacturaDespensa.Caption = FacturaDespensa.Caption & ": " & UCase(FA.NombreEstab)
            CodigoL = FA.Cod_CxC
            Cta_Cobrar = FA.Cta_CxP
            Ln_No = 1
            Cant_Item_PV = 100
            TextCant.Text = "0"
            TextVUnit.Text = "0"
            LabelVTotal.Caption = "0"
            Modificar = False
            Bandera = True
            Label1.Caption = "DESPENSA No. " & FA.Autorizacion & "-" & FA.Serie & "-"
            FA.Nuevo_Doc = True
            FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
            TextFacturaNo = Format(FA.Factura, "000000000")
                  
            sSQL = "DELETE * " _
                 & "FROM Asiento_F " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND CodigoU = '" & CodigoUsuario & "' "
            Ejecutar_SQL_SP sSQL
        
            sSQL = "SELECT CodBod " _
                 & "FROM Catalogo_Bodegas " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND CodBod = '" & Cod_Bodega & "' "
            SelectDB_Combo DCBodega, AdoBodega, sSQL, "CodBod"
            Cod_Bodega = Ninguno
            If AdoBodega.Recordset.RecordCount > 0 Then
               Cod_Bodega = AdoBodega.Recordset.fields("CodBod")
               sSQL = "SELECT CP.Producto, CP.Codigo_Inv, CP.Codigo_Barra, SUM(TK.Entrada-TK.Salida) " _
                    & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
                    & "WHERE CP.Item = '" & NumEmpresa & "' " _
                    & "AND CP.Periodo = '" & Periodo_Contable & "' " _
                    & "AND TK.CodBodega = '" & Cod_Bodega & "' " _
                    & "AND CP.TC = 'P' " _
                    & "AND CP.Item = TK.Item " _
                    & "AND CP.Periodo = TK.Periodo " _
                    & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
                    & "GROUP BY TK.CodBodega, CP.Producto, CP.Codigo_Inv, CP.Codigo_Barra " _
                    & "HAVING SUM(TK.Entrada - TK.Salida) > 0 " _
                    & "ORDER BY CP.Producto, CP.Codigo_Inv "
               SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
               RatonNormal
               If AdoArticulo.Recordset.RecordCount > 0 Then
                  FacturaDespensa.WindowState = 2
                  FA.Fecha_Desde = BuscarFecha(PrimerDiaMes(MBFecha))
                  FA.Fecha_Hasta = BuscarFecha(UltimoDiaMes(MBFecha))
              '   DGAsientoF.width = MDI_X_Max - 100
                  DGAsientoF.Height = MDI_Y_Max - DGAsientoF.Top - 1000
                  DGAsientoF.Refresh
                  Label26.Top = DGAsientoF.Top + DGAsientoF.Height + 100
                  LabelTotal.Top = DGAsientoF.Top + Label26.Height + DGAsientoF.Height + 60
                  Command1.Top = DGAsientoF.Top + DGAsientoF.Height + 100
                  Command3.Top = DGAsientoF.Top + DGAsientoF.Height + 100
                  sSQL = "SELECT * " _
                       & "FROM Asiento_F " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND CodigoU = '" & CodigoUsuario & "' "
                  SQLDec = "PRECIO 4|CORTE 5|."
                  Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
                  Buscar_Cliente Ninguno
               Else
                  MsgBox "No existen Productos a despachar"
                  Unload FacturaDespensa
               End If
            Else
               MsgBox "No existe Bodega asignada"
               Unload FacturaDespensa
            End If
         Else
            MsgBox "Falta Organizar la CxC en Puntos de Venta." & vbCrLf _
                 & "Salga de este proceso y llame al su técnico" & vbCrLf _
                 & "o al Contador de su Organizacion."
            Unload FacturaDespensa
         End If
     Else
        MsgBox "Error: Falta asignar la bodega a la despensa, llame al su técnico o al Contador de su Organizacion."
        Unload FacturaDespensa
     End If
  Else
    Cod_Bodega = Ninguno
    SerieFactura = "000000"
    MsgBox "Falta Organizar la Despensa." & vbCrLf _
         & "Salga de este proceso y llame al su técnico" & vbCrLf _
         & "o al Contador de su Organizacion."
    Unload FacturaDespensa
  End If
End Sub

Private Sub Form_Deactivate()
  FacturaDespensa.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCliente
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFactura
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  FechaTexto1 = MBFecha
  FA.Fecha = MBFecha
  FA.Fecha_V = MBFecha
  FA.Fecha_Desde = BuscarFecha(PrimerDiaMes(MBFecha))
  FA.Fecha_Hasta = BuscarFecha(UltimoDiaMes(MBFecha))
End Sub

Private Sub TextCant_GotFocus()
  If Val(TextVUnit) <= 0 Then TextVUnit = Format(Precio, "#,##0.00")
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
Dim Grabar_PV As Boolean

  If Cod_Bodega = "" Then Cod_Bodega = "01"
  If Cod_Bodega = "." Then Cod_Bodega = "01"
  
  TextoValido TextVUnit, True, , 2
  TextCant = Val(TextCant)
  TextVUnit = Val(TextVUnit)
  If IsNumeric(TextVUnit) And IsNumeric(TextCant) Then
     If Val(TextVUnit) = 0 Then TextVUnit = "0.01"
     If Val(TextCant) = 0 Then TextCant = "1"
  Else
     If Val(TextVUnit) = 0 Then TextVUnit = "0.01"
     If Val(TextCant) = 0 Then TextCant = "1"
  End If
   Grabar_PV = True
   If Cant_Item_PV > 0 And (AdoAsientoF.Recordset.RecordCount > Cant_Item_PV) Then Grabar_PV = False
   'MsgBox Cant_Item_PV
   If Grabar_PV Then
      Real1 = 0: Real2 = 0: Real3 = 0
      'Real1 = Redondear(CCur(TextCant) * CCur(TextVUnit), 2)
      With AdoAsientoF.Recordset
        If Real1 >= 0 Then
           SetAddNew AdoAsientoF
           SetFields AdoAsientoF, "CODIGO", Codigos
           SetFields AdoAsientoF, "CODIGO_L", CodigoL
           SetFields AdoAsientoF, "PRODUCTO", Mid(Producto, 1, 45)
           SetFields AdoAsientoF, "CANT", CCur(TextCant)
           SetFields AdoAsientoF, "PRECIO", 0 ' CCur(TextVUnit)
           SetFields AdoAsientoF, "CodBod", Cod_Bodega
           SetFields AdoAsientoF, "TOTAL", Real1
           SetFields AdoAsientoF, "Total_IVA", Real3
           SetFields AdoAsientoF, "COD_BAR", DatInv.Codigo_Barra
           SetFields AdoAsientoF, "COSTO", DatInv.Costo
           SetFields AdoAsientoF, "Cta_Inv", DatInv.Cta_Inventario
           SetFields AdoAsientoF, "Cta_Costo", DatInv.Cta_Costo_Venta
           SetFields AdoAsientoF, "Item", NumEmpresa
           SetFields AdoAsientoF, "CodigoU", CodigoUsuario
           SetFields AdoAsientoF, "A_No", Ln_No
           SetUpdate AdoAsientoF
           Calculos_Totales_Factura AdoAsientoF
           Ln_No = Ln_No + 1
           TextVUnit.Text = ""
        End If
      End With
      LabelVTotal.Caption = Format(Real1, "#,##0.00")
   Else
      MsgBox "Ya no puede ingresar mas productos"
   End If
  DCArticulo.SetFocus
  
End Sub

Public Sub Buscar_Cliente(Busqueda As String)
    sSQL = "SELECT DISTINCT TOP 50 C.Cliente, C.CI_RUC, C.Codigo, C.Cta_CxP, C.Grupo, C.Cod_Ejec, F.Cont_Salidas " _
         & "FROM Clientes As C, Facturas As F " _
         & "WHERE F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.TC IN ('NV','FA') " _
         & "AND F.T <> 'A' " _
         & "AND F.Cont_Salidas > 0 " _
         & "AND F.Fecha BETWEEN #" & FA.Fecha_Desde & "# AND #" & FA.Fecha_Hasta & "# " _
         & "AND C.Grupo = '" & MidStrg(FA.Serie, 4, 3) & "' "
    If Len(Busqueda) > 1 Then
       If IsNumeric(Busqueda) Then
          sSQL = sSQL & "AND C.CI_RUC LIKE '" & Busqueda & "%' "
       Else
          sSQL = sSQL & "AND C.Cliente LIKE '%" & Busqueda & "%' "
       End If
    End If
    sSQL = sSQL _
         & "AND C.Codigo=F.CodigoC " _
         & "ORDER BY Cliente "
   'MsgBox sSQL
    SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
End Sub

Public Sub ProcGrabarDespensa()
 'Seteamos los encabezados para las facturas
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     DGAsientoF.Visible = False
     Codigo1 = DCCliente.Text
     If Codigo1 = "" Then Codigo1 = Ninguno
     TBeneficiario.Patron_Busqueda = Codigo1
     TBeneficiario = Leer_Datos_Cliente_SP(Codigo1)
     If TBeneficiario.Codigo <> Ninguno Then
        CodigoCliente = TBeneficiario.Codigo
        NombreCliente = TBeneficiario.Cliente
        CICliente = TBeneficiario.CI_RUC
        Grupo_No = TBeneficiario.Grupo_No
        TipoDoc = TBeneficiario.TD
        FA.CodigoC = CodigoCliente
        FA.Cliente = NombreCliente
        FA.CI_RUC = CICliente
        CmdCIBenef1.Caption = TBeneficiario.CI_RUC
     
        FechaTexto = MBFecha
        FA.Fecha = MBFecha
        FA.Fecha_V = MBFecha
        HoraTexto = Format(Time, FormatoTimes)
        Calculos_Totales_Factura AdoAsientoF
        FA.T = Pendiente
        FA.SubTotal = Total_Sin_IVA + Total_Con_IVA + Total_Servicio - Total_Desc
        FA.Descuento = Total_Desc
        FA.Servicio = Total_Servicio
        FA.Con_IVA = Total_Con_IVA
        FA.Sin_IVA = Total_Sin_IVA
        FA.Total_IVA = Total_IVA
        FA.Total_MN = Total_Factura
        FA.Saldo_MN = Total_Factura
     
       'Datos del Representante
        FA.Razon_Social = TBeneficiario.Representante
        FA.RUC_CI = TBeneficiario.RUC_CI_Rep
        FA.TB = TBeneficiario.TD_Rep
        FA.TelefonoC = TBeneficiario.Telefono1
        FA.DireccionC = TBeneficiario.Direccion_Rep
        FA.EmailR = TBeneficiario.EmailR
     
       'Grabamos el numero de factura
        FA.Nuevo_Doc = True
        FA.Tipo_PRN = "FM"
        Grabar_Factura FA, True
        
        sSQL = "UPDATE Facturas " _
             & "SET Cont_Salidas = Cont_Salidas - 1 " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC IN ('NV','FA') " _
             & "AND T <> 'A' " _
             & "AND Cont_Salidas > 0 " _
             & "AND CodigoC = '" & FA.CodigoC & "' " _
             & "AND Fecha BETWEEN #" & FA.Fecha_Desde & "# AND #" & FA.Fecha_Hasta & "# "
        Ejecutar_SQL_SP sSQL
        
       'Abono de Factura
'''        TA.T = Normal
'''        TA.TP = FA.TC
'''        TA.Serie = FA.Serie
'''        TA.Autorizacion = FA.Autorizacion
'''        TA.CodigoC = FA.CodigoC
'''        TA.Factura = FA.Factura
'''        TA.Fecha = FA.Fecha
'''        TA.Cta = Cta_CajaG
'''        TA.Cta_CxP = FA.Cta_CxP
'''        TA.Banco = "ABONO POR VERIFICAR"
'''        TA.Cheque = UCase$(Grupo_No)
'''        TA.Abono = FA.Total_MN
'''        Grabar_Abonos TA
    
       'Autorizamos la factura i/o Guia de Remision
        RatonNormal
        Ln_No = 1
        Imprimir_Punto_Venta FA
'        Imprimir_Facturas_CxC FacturaDespensa, FA
     End If
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub Datos_Articulos_Despensa()
    Codigos = DatInv.Codigo_Inv
    Producto = DatInv.Producto
    Cta_Ventas = DatInv.Cta_Ventas
    Precio = 0 'Redondear(DatInv.PVP, 2)
    DatInv.PVP2 = Redondear(DatInv.PVP2, 2)
    LabelStock.Caption = DatInv.Stock
    BanIVA = DatInv.IVA
    TextVUnit = Format(Precio, "#,##0.0000")
    If TipoFactura = "NV" Then BanIVA = False
    DCArticulo.Text = Producto
End Sub

