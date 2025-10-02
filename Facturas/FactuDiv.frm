VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturasDiv 
   Caption         =   "FACTURACION/LIQUIDACION DE COMPRAS"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   10425
   WindowState     =   1  'Minimized
   Begin VB.PictureBox PictBarra 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   8925
      ScaleHeight     =   585
      ScaleWidth      =   1950
      TabIndex        =   37
      Top             =   105
      Width           =   2010
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
      Height          =   4740
      Left            =   1785
      TabIndex        =   35
      Top             =   210
      Visible         =   0   'False
      Width           =   7365
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "FactuDiv.frx":0000
         DataSource      =   "AdoBenef"
         Height          =   4275
         Left            =   105
         TabIndex        =   36
         Top             =   210
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   7541
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
   Begin VB.TextBox TxtVTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   450
      Left            =   105
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "FactuDiv.frx":0017
      Top             =   2415
      Width           =   2115
   End
   Begin VB.CommandButton Command5 
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
      Left            =   7560
      Picture         =   "FactuDiv.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2100
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Calcular"
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
      Left            =   6195
      Picture         =   "FactuDiv.frx":08E5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2100
      Width           =   1275
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
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "CONSUMIDOR FINAL"
      ToolTipText     =   "Escriba el Nombre o C.I./RUC del Beneficiario o las primeras letras del Apellido"
      Top             =   420
      Width           =   7365
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
      Left            =   7875
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "00000000"
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox TextVUnit 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   450
      Left            =   6930
      MaxLength       =   13
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "FactuDiv.frx":0BEF
      Top             =   1575
      Width           =   1905
   End
   Begin VB.TextBox TextCant 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   450
      Left            =   2310
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "FactuDiv.frx":0BF6
      Top             =   2415
      Width           =   2115
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FactuDiv.frx":0BFA
      DataSource      =   "AdoArticulo"
      Height          =   420
      Left            =   105
      TabIndex        =   9
      Top             =   1575
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   741
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin VB.TextBox TxtEfectivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   420
      Left            =   7035
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "FactuDiv.frx":0C14
      Top             =   6090
      Width           =   2010
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Grabar"
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
      Left            =   9135
      Picture         =   "FactuDiv.frx":0C1B
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5040
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9135
      Picture         =   "FactuDiv.frx":0F25
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5880
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "FactuDiv.frx":17EF
      Height          =   2010
      Left            =   105
      TabIndex        =   18
      Top             =   2940
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   3545
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   13
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
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
      Left            =   840
      Top             =   3570
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
      Left            =   840
      Top             =   3885
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
      Left            =   840
      Top             =   4200
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
      Left            =   840
      Top             =   3570
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
      Left            =   2940
      Top             =   3570
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
      Left            =   2940
      Top             =   3885
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
      Left            =   2940
      Top             =   4200
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
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "FactuDiv.frx":1809
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   1995
      TabIndex        =   5
      Top             =   840
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Clientes"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nota de Venta No."
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
      Left            =   5355
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE PROCESO"
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
      TabIndex        =   4
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL EN"
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
      TabIndex        =   12
      Top             =   2100
      Width           =   2115
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
      Left            =   6930
      TabIndex        =   10
      Top             =   1260
      Width           =   1905
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad Cambio"
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
      Left            =   2310
      TabIndex        =   14
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label LabelTotalME 
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
      Left            =   7035
      TabIndex        =   28
      Top             =   5565
      Width           =   2010
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Fact. (ME)"
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
      Left            =   4620
      TabIndex        =   27
      Top             =   5565
      Width           =   2430
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
      Left            =   1470
      TabIndex        =   2
      Top             =   105
      Width           =   7365
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
      TabIndex        =   8
      Top             =   1260
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label LblCambio 
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
      Left            =   7035
      TabIndex        =   32
      Top             =   6615
      Width           =   2010
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EFECTIVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4620
      TabIndex        =   29
      Top             =   6090
      Width           =   2430
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAMBIO"
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
      Left            =   4620
      TabIndex        =   31
      Top             =   6615
      Width           =   2430
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
      Left            =   7035
      TabIndex        =   26
      Top             =   5040
      Width           =   2010
   End
   Begin VB.Label LabelIVA 
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
      Left            =   2520
      TabIndex        =   24
      Top             =   6090
      Width           =   2010
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado"
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
      Left            =   4620
      TabIndex        =   25
      Top             =   5040
      Width           =   2430
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A. 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   105
      TabIndex        =   23
      Top             =   6090
      Width           =   2430
   End
   Begin VB.Label LabelConIVA 
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
      Left            =   2520
      TabIndex        =   22
      Top             =   5565
      Width           =   2010
   End
   Begin VB.Label LabelSubTotal 
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
      Left            =   2520
      TabIndex        =   20
      Top             =   5040
      Width           =   2010
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 12%"
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
      Left            =   105
      TabIndex        =   21
      Top             =   5565
      Width           =   2430
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 0%"
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
      Left            =   105
      TabIndex        =   19
      Top             =   5040
      Width           =   2430
   End
End
Attribute VB_Name = "FacturasDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Grupo_Inv As String

Private Sub Command2_Click()
'Opcion Multiplicar
   TextoValido TextCant, True, , 3
   TextoValido TextVUnit, True, , 4
   TextoValido TxtVTotal, True, , 2
   If IsNumeric(TextVUnit) And IsNumeric(TextCant) And IsNumeric(TxtVTotal) Then
      If CCur(TextVUnit) <= 0 Then TextVUnit = "1"
      
      If CCur(TxtVTotal) > 0 And CCur(TextCant) = 0 Then
         If DatInv.Div Then
            TextCant = Format$(CCur(TxtVTotal) / CCur(TextVUnit), "#,##0.000")
         Else
            TextCant = Format$(CCur(TxtVTotal) * CCur(TextVUnit), "#,##0.000")
         End If
      ElseIf CCur(TextCant) > 0 And CCur(TxtVTotal) = 0 Then
         If DatInv.Div Then
            TxtVTotal = Format$(CCur(TextCant) / CCur(TextVUnit), "#,##0.00")
         Else
            TxtVTotal = Format$(CCur(TextCant) * CCur(TextVUnit), "#,##0.00")
         End If
      End If
   End If
End Sub

Private Sub Command5_Click()
Dim Grabar_PV As Boolean
   Grabar_PV = True
   If Cant_Item_PV > 0 And (AdoAsientoF.Recordset.RecordCount > Cant_Item_PV) Then Grabar_PV = False
   'MsgBox Cant_Item_PV
   If Grabar_PV Then
      TextoValido TextCant, True, , 2
      TextoValido TextVUnit, True, , 4
      TextoValido TxtVTotal, True, , 2
      Real1 = CCur(TxtVTotal.Text): Real2 = 0: Real3 = 0
      With AdoAsientoF.Recordset
        If Real1 > 0 Then
           If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
           SetAddNew AdoAsientoF
           SetFields AdoAsientoF, "CODIGO", Codigos
           SetFields AdoAsientoF, "CODIGO_L", CodigoL
           SetFields AdoAsientoF, "PRODUCTO", MidStrg(Producto, 1, 120)
           SetFields AdoAsientoF, "CANT", CCur(TextCant)
           SetFields AdoAsientoF, "PRECIO", CCur(TextVUnit)
           SetFields AdoAsientoF, "TOTAL", Real1
           SetFields AdoAsientoF, "Total_IVA", Real3
           SetFields AdoAsientoF, "Item", NumEmpresa
           SetFields AdoAsientoF, "CodigoU", CodigoUsuario
           SetFields AdoAsientoF, "A_No", Ln_No
           SetUpdate AdoAsientoF
           Calculos_Totales_Factura FA
           Ln_No = Ln_No + 1
        End If
      End With
      
      sSQL = "UPDATE Catalogo_Productos " _
           & "SET PVP = " & CCur(TextVUnit) & " " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Codigo_Inv = '" & DatInv.Codigo_Inv & "' "
      Ejecutar_SQL_SP sSQL
      
      If FA.TC = "LC" Then Grupo_Inv = "02" Else Grupo_Inv = "01"
      sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
           & "FROM Catalogo_Productos " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = 'P' " _
           & "AND MidStrg(Codigo_Inv,1,2) = '" & Grupo_Inv & "' " _
           & "ORDER BY Producto,Codigo_Inv "
      SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
      TextCant = "0.00"
      TxtVTotal = "0.00"
   Else
      MsgBox "Ya no puede ingresar mas productos"
   End If
  DCArticulo.SetFocus
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
 If TextFacturaNo = "" Then TextFacturaNo = "0"
 Factura_No = Val(TextFacturaNo)
 FA.Factura = Val(TextFacturaNo)
 FA.Nuevo_Doc = False
 If Existe_Factura(FA) Then
    FA.Factura = TextFacturaNo
 Else
    If Val(TextFacturaNo) < ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False) Then
       Mensajes = "La " & FA.TC & " No. " & TextFacturaNo & vbCrLf _
                & "No esta Procesada, Desea Procesarla"
       Titulo = "Formulario de Confirmación"
       If BoxMensaje = vbYes Then
          FA.Factura = TextFacturaNo
          FA.Nuevo_Doc = False
       End If
    Else
       FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
       TextFacturaNo = FA.Factura
       FA.Nuevo_Doc = True
    End If
 End If
 DCArticulo.SetFocus
End Sub

Private Sub DCLinea_GotFocus()
  Grupo_No = Ninguno
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  FA.Cod_CxC = DCLinea
  FA.Fecha = MBFecha
  Lineas_De_CxC FA
  Tipo_De_Facturacion
  TipoFactura = FA.TC
  If FA.TC = "LC" Then Grupo_Inv = "03.02" Else Grupo_Inv = "03.01"
  sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND MidStrg(Codigo_Inv,1,5) = '" & Grupo_Inv & "' " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  NumComp = FA.Factura
  TextFacturaNo.Text = Format$(NumComp, "0000000")
End Sub

Public Sub Tipo_De_Facturacion()
  If FA.TC = "NV" Then
     FacturasDiv.Caption = "INGRESAR NOTA DE VENTA"
     Label2.Caption = "(" & FA.Autorizacion & ") No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. 0.00%"
  ElseIf FA.TC = "OP" Then
     FacturasDiv.Caption = "INGRESAR ORDEN DE PEDIDO"
     Label2.Caption = "(" & FA.Autorizacion & ") No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. 0.00%"
  ElseIf FA.TC = "LC" Then
     FacturasDiv.Caption = "LIQUIDACION DE COMPRAS"
     Label2.Caption = "(" & FA.Autorizacion & ") No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
  Else
     FacturasDiv.Caption = "INGRESAR FACTURAS"
     Label2.Caption = "(" & FA.Autorizacion & ") No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
  End If
  TipoFactura = FA.TC
  'FA.Factura = Numero_Factura(FA)
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
          CodigoBenef = .fields("Codigo")
          TextBenef.Text = .fields("Cliente")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          FrmBenef.Visible = False
          'TextCotiza.SetFocus
       Else
          NombreCliente = DCCliente.Text
          FrmBenef.Visible = False
          FacturasDiv.Visible = False
          Nuevo = True
          MsgBox "Cliente No existe"
          FClientesFlash.Show 1
          FacturasDiv.Visible = True
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

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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
      .Find ("Cliente = '" & TextBenef.Text & "'")
       If Not .EOF Then
          CodigoBenef = .fields("Codigo")
          TextBenef.Text = .fields("Cliente")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          FA.CodigoC = CodigoCliente
          FA.Cliente = NombreCliente
       Else
          DCCliente.Text = TextBenef.Text
          FrmBenef.Visible = True
          DCCliente.SetFocus
       End If
   Else
       DCCliente.Text = TextBenef.Text
       FrmBenef.Visible = True
       DCCliente.SetFocus
   End If
  End With
End Sub

Public Sub Listar_Clientes()
  sSQL = "SELECT TOP 100 Cliente,Codigo,CI_RUC,TD,Grupo " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND FA <> " & Val(adFalse) & " " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
End Sub

''Public Sub Calculos_Totales_Factura(DtaProd As Adodc)
''Dim NumLn As Byte
''  Total_ME = 0
''  Si_No = False
''  Total_Factura = 0: Total_Desc = 0
''  Total_Con_IVA = 0: Total_Sin_IVA = 0
''  Total_IVA = 0
''  NumLn = 0
''  With DtaProd.Recordset
''   If .RecordCount > 0 Then
''      .MoveFirst
''       Do While Not .EOF
''          Total_IVA = Total_IVA + .Fields("Total_IVA")
''          Total_Desc = Total_Desc + .Fields("Total_Desc")
''          If .Fields("Total_IVA") > 0 Then
''             Total_Con_IVA = Total_Con_IVA + .Fields("TOTAL")
''          Else
''             Total_Sin_IVA = Total_Sin_IVA + .Fields("TOTAL")
''          End If
''          NumLn = NumLn + 1
''         .MoveNext
''       Loop
''   End If
''  End With
''  Total_Con_IVA = Redondear(Total_Con_IVA, 2)
''  Total_Sin_IVA = Redondear(Total_Sin_IVA, 2)
''  Total_Servicio = Redondear((Total_Sin_IVA + Total_Con_IVA - Total_Desc) * Porc_Serv, 2)
''  Total_Factura = Redondear(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
''  LabelSubTotal.Caption = Format$(Total_Sin_IVA, "#,##0.00")
''  LabelConIVA.Caption = Format$(Total_Con_IVA, "#,##0.00")
''  LabelIVA.Caption = Format$(Total_IVA, "#,##0.00")
''  LabelTotal.Caption = Format$(Total_Factura, "#,##0.00")
''  Total_FacturaME = 0
''  If Val(TextCotiza) > 0 Then
''     TotalDolar = Val(CCur(TextCotiza))
''     If OpcDiv.value Then
''        Total_FacturaME = Redondear(Total_Factura / TotalDolar, 2)
''     Else
''        Total_FacturaME = Redondear(Total_Factura * TotalDolar, 2)
''     End If
''  End If
''  LabelTotalME.Caption = Format$(Total_FacturaME, "#,##0.00")
''  Label7.Caption = " CAMBIO (" & NumLn & ")"
''  TextCant.Text = ""
''  TxtVTotal = ""
''End Sub

Private Sub Command1_Click()
  Unload FacturasDiv
End Sub

Private Sub Command3_Click()
  FechaValida MBFecha
  FechaTexto = MBFecha
  Calculos_Totales_Factura FA
  If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then
  If FA.TC = "LC" Then Grupo_Inv = "02" Else Grupo_Inv = "01"
     Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
              & FA.TC & " No. " & FA.Factura
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = vbYes Then
        Moneda_US = False
        TextoFormaPago = PagoCont
        ProcGrabar
        
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL
        Ln_No = 1
        DCArticulo.SetFocus
     End If
  Else
     MsgBox "Error: El Efectivo no alcanza para grabar"
  End If
End Sub

Private Sub DCArticulo_GotFocus()
  Calculos_Totales_Factura FA
    LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
    LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
    LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
    LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
    TxtEfectivo = Format(FA.Total_MN, "#,##0.00")
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
   Select Case KeyCode
     Case vbKeyEscape
          Calculos_Totales_Factura FA
          'TxtEfectivo = "0.00"
          TxtEfectivo.SetFocus
     Case vbKeyReturn
          TextCant.SetFocus
   End Select
  If KeyCode = vbKeyF1 Then
     With AdoArticulo.Recordset
      If .RecordCount Then
         .MoveFirst
         .Find ("Producto = '" & DCArticulo & "' ")
          If Not .EOF Then MsgBox .fields("Producto") & ":" & vbCrLf & .fields("Ayuda")
      End If
     End With
  End If
End Sub

Private Sub DCArticulo_LostFocus()
  Codigos = Ninguno
  If Leer_Codigo_Inv(DCArticulo, FechaSistema, Cod_Bodega) Then DatosArticulos
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
  Cod_Bodega = Ninguno
  Label20.Caption = "TOTAL EN " & Moneda
  FA.TC = TipoFactura
  FA.Fecha = MBFecha
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
       & "ORDER BY Fact,Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto", , "Catalogo Lineas"
  FA.Cod_CxC = DCLinea
  Cant_Item_PV = 30
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  CodigoCliente = "9999999999"
  NombreCliente = "CONSUMIDOR FINAL"
  DireccionCli = " S/N"
  TextBenef = NombreCliente
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  TextCant.Text = "0"
  TextVUnit.Text = "0.00"
  TxtVTotal.Text = "0.00"
  Modificar = False
  Bandera = True
  Mifecha = BuscarFecha(FechaSistema)
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  SerieFactura = "999999"
  Autorizacion = "9999999999"
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SQLDec = "PRECIO 3|"
  Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
  Listar_Clientes
  MBFecha.Text = FechaSistema
  FacturasDiv.WindowState = 2
  RatonNormal
End Sub

Private Sub Form_Deactivate()
  FacturasDiv.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoBenef
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFactura
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
  
  SRI_Obtener_Datos_Comprobantes_Electronicos
  
  MBFecha.Text = FechaSistema
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  
  FechaTexto1 = MBFecha.Text
  FA.Fecha = MBFecha.Text
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
       & "ORDER BY Fact,Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True, , 2
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True, , 4
End Sub

Public Sub ProcGrabar()
 DGAsientoF.Visible = False
 TipoFactura = FA.TC
 FA.Porc_IVA = Porc_IVA
 FA.T = Pendiente
 FA.Forma_Pago = PagoCred
 'Seteamos los encabezados para las facturas
  Calculos_Totales_Factura FA
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     FechaTexto = MBFecha.Text
     HoraTexto = Format$(Time, FormatoTimes)
     Total_FacturaME = 0
     Moneda_US = False
     If FA.Nuevo_Doc Then FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
     If TipoFactura = "PV" Then
        Control_Procesos "F", "Grabar Ticket No. " & FA.Factura
     ElseIf TipoFactura = "NV" Then
        Control_Procesos "F", "Grabar Nota de Venta No. " & FA.Factura
     ElseIf TipoFactura = "CP" Then
        Control_Procesos "F", "Grabar Cheque Protestado No. " & FA.Factura
     ElseIf TipoFactura = "LC" Then
        Control_Procesos "F", "Grabar Liquidacion de Compras No. " & FA.Factura
     Else
        Control_Procesos "F", "Grabar Factura No. " & FA.Factura
     End If
     If Moneda_US Then Total = FA.Total_ME Else Total = FA.Total_MN
     FA.Efectivo = Val(CCur(TxtEfectivo))
     Dolar = CCur(Val(TextCotiza))
     FA.Cotizacion = Dolar
    
    'Grabamos el numero de factura
     Grabar_Factura FA, True
     
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL
     RatonNormal
     If TipoFactura <> "CP" Then
        Evaluar = True
        FechaTexto = MBFecha.Text
        TA.T = Normal
        TA.TP = FA.TC
        TA.Fecha = MBFecha
        TA.Cta = Cta_CajaG
        TA.Banco = "EFECTIVO MN"
        TA.Cheque = Format$(Factura_No, "00000000")
        TA.Factura = FA.Factura
        TA.Abono = Total_Factura
        TA.Serie = FA.Serie
        TA.Autorizacion = FA.Autorizacion
        TA.CodigoC = FA.CodigoC
        Grabar_Abonos TA
        sSQL = "UPDATE Facturas " _
             & "SET Saldo_MN = 0, T = 'C' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Factura = " & FA.Factura & " " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND CodigoC = '" & CodigoCliente & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND Serie = '" & FA.Serie & "' "
        Ejecutar_SQL_SP sSQL
     End If
     Imprimir_Punto_Venta FA
     NombreCliente = "CONSUMIDOR FINAL"
     TextBenef = "CONSUMIDOR FINAL"
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub DatosArticulos()
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
       Codigos = DatInv.Codigo_Inv
       Producto = DatInv.Producto
       Cta_Ventas = DatInv.Cta_Ventas
       Precio = DatInv.PVP
       BanIVA = DatInv.IVA
       TextVUnit = Format$(Precio, "#,##0.0000")
       If TipoFactura = "NV" Then BanIVA = False
       DCArticulo.Text = Producto
   End If
  End With
End Sub

Private Sub TxtEfectivo_GotFocus()
  MarcarTexto TxtEfectivo
End Sub

Private Sub TxtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_Change()
'Total_FacturaME
  If IsNumeric(TxtEfectivo) Then
     If Val(TextCotiza) > 0 Then
        If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - FA.Total_ME, "#,##0.00")
     Else
        If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - FA.Total_MN, "#,##0.00")
     End If
  End If
End Sub

Private Sub TxtEfectivo_LostFocus()
  TextoValido TxtEfectivo, True, , 2
  If Val(TextCotiza) > 0 Then
     LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - FA.Total_ME, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - FA.Total_ME) >= 0 Then Command3.SetFocus
  Else
     LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - FA.Total_MN, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - FA.Total_MN) >= 0 Then Command3.SetFocus
  End If
End Sub

Private Sub TxtVTotal_GotFocus()
  MarcarTexto TxtVTotal
End Sub

Private Sub TxtVTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtVTotal_LostFocus()
  TextoValido TxtVTotal, True, , 2
End Sub
