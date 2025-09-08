VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturaFarmacia 
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   15960
   WindowState     =   1  'Minimized
   Begin VB.CommandButton CmdCIBenef3 
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
      Left            =   9450
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1365
      Width           =   1695
   End
   Begin VB.CommandButton CmdCIBenef2 
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
      Left            =   9450
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   945
      Width           =   1695
   End
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
      Left            =   9450
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   525
      Width           =   1695
   End
   Begin VB.TextBox TxtBeneficiario2 
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
      Left            =   1785
      TabIndex        =   11
      Top             =   1365
      Width           =   7575
   End
   Begin VB.TextBox TxtBeneficiario 
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
      Left            =   1785
      TabIndex        =   8
      Top             =   945
      Width           =   7575
   End
   Begin VB.TextBox TxtAfiliado 
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
      Left            =   1785
      TabIndex        =   5
      Top             =   525
      Width           =   7575
   End
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   9135
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   44
      Top             =   6195
      Width           =   1200
   End
   Begin VB.Frame FrmGrupo 
      BackColor       =   &H00400000&
      Caption         =   "| GRUPO DE FACTURACION |"
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
      Left            =   1575
      TabIndex        =   42
      Top             =   2205
      Visible         =   0   'False
      Width           =   5895
      Begin MSDataListLib.DataList DLGrupo 
         Bindings        =   "FacturaF.frx":0000
         DataSource      =   "AdoGrupo"
         Height          =   2940
         Left            =   105
         TabIndex        =   43
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   1785
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
      Left            =   8400
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "FacturaF.frx":0017
      Top             =   2625
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
      Left            =   7455
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "FacturaF.frx":001C
      Top             =   2625
      Width           =   960
   End
   Begin VB.TextBox TxtDocumentos 
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
      Left            =   6090
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2625
      Width           =   1380
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "FacturaF.frx":001E
      DataSource      =   "AdoBodega"
      Height          =   420
      Left            =   1575
      TabIndex        =   14
      Top             =   1785
      Width           =   5685
      _ExtentX        =   10028
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9345
      TabIndex        =   3
      Text            =   "0"
      Top             =   105
      Width           =   2115
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FacturaF.frx":0036
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   18
      Top             =   2625
      Width           =   6000
      _ExtentX        =   10583
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
      TabIndex        =   39
      Text            =   "FacturaF.frx":0050
      Top             =   7245
      Width           =   2010
   End
   Begin VB.CommandButton Command3 
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
      Left            =   10290
      Picture         =   "FacturaF.frx":0057
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6195
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
      Left            =   10290
      Picture         =   "FacturaF.frx":0361
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7035
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "FacturaF.frx":0C2B
      Height          =   3060
      Left            =   105
      TabIndex        =   27
      Top             =   3045
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   5398
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
      Left            =   525
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
      Top             =   4515
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
      Top             =   3885
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
      Top             =   4830
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
      Top             =   5145
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
      Left            =   2625
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   2625
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
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO 2"
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
      TabIndex        =   10
      Top             =   1365
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   945
      Width           =   1695
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
      Left            =   9765
      TabIndex        =   26
      Top             =   2625
      Width           =   1695
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
      Left            =   9765
      TabIndex        =   25
      Top             =   2310
      Width           =   1695
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
      Left            =   8400
      TabIndex        =   23
      Top             =   2310
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
      Left            =   7455
      TabIndex        =   21
      Top             =   2310
      Width           =   960
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
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
      TabIndex        =   19
      Top             =   2310
      Width           =   1380
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
      Left            =   9345
      TabIndex        =   16
      Top             =   1785
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
      Left            =   7350
      TabIndex        =   15
      Top             =   1785
      Width           =   2010
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
      TabIndex        =   37
      Top             =   6720
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
      TabIndex        =   36
      Top             =   6720
      Width           =   2430
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " AFILIADO"
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
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   17
      Top             =   2310
      Width           =   6000
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
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LINEA:"
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
      Left            =   3150
      TabIndex        =   2
      Top             =   105
      Width           =   6210
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
      Width           =   1695
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
      TabIndex        =   38
      Top             =   7245
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
      TabIndex        =   35
      Top             =   6195
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
      TabIndex        =   33
      Top             =   7245
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
      TabIndex        =   34
      Top             =   6195
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
      TabIndex        =   32
      Top             =   7245
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
      TabIndex        =   31
      Top             =   6720
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
      TabIndex        =   29
      Top             =   6195
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
      TabIndex        =   30
      Top             =   6720
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
      TabIndex        =   28
      Top             =   6195
      Width           =   2430
   End
End
Attribute VB_Name = "FacturaFarmacia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Grupo_Inv As String

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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
  Total_Con_IVA = Redondear(Total_Con_IVA, 2)
  Total_Sin_IVA = Redondear(Total_Sin_IVA, 2)
  Total_Servicio = Redondear((Total_Sin_IVA + Total_Con_IVA - Total_Desc) * Porc_Serv, 2)
  Total_Factura = Redondear(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
  LabelSubTotal.Caption = Format(Total_Sin_IVA, "#,##0.00")
  LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0.00")
  LabelIVA.Caption = Format(Total_IVA, "#,##0.00")
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
  Total_FacturaME = 0
  If Val(TextCotiza) > 0 Then
     TotalDolar = Val(CCur(TextCotiza))
     If OpcDiv.value Then
        Total_FacturaME = Redondear(Total_Factura / TotalDolar, 2)
     Else
        Total_FacturaME = Redondear(Total_Factura * TotalDolar, 2)
     End If
  End If
  LabelTotalME.Caption = Format(Total_FacturaME, "#,##0.00")
  TxtEfectivo = Format(Total_Factura, "#,##0.00")
  TextCant.Text = ""
  LabelVTotal.Caption = ""
End Sub

Private Sub Command1_Click()
  Unload FacturaFarmacia
End Sub

Private Sub Command3_Click()
    FechaValida MBFecha
    FechaTexto = MBFecha
    FA.Fecha = MBFecha
    FA.Fecha_V = MBFecha
    Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
             & "La Factura No. " & TextFacturaNo.Text
    Titulo = "Formulario de Grabacion"
    If BoxMensaje = vbYes Then
       Moneda_US = False
       TextoFormaPago = PagoCont
       ProcGrabar
       FA.Nuevo_Doc = True
       FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
       TextFacturaNo = FA.Factura
       
       sSQL = "DELETE * " _
            & "FROM Asiento_F " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       Ejecutar_SQL_SP sSQL
       
       sSQL = "SELECT * " _
            & "FROM Asiento_F " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL
       
       FacturaFarmacia.Caption = "INGRESAR FACTURA"
       Label1.Caption = FA.Autorizacion & " - FACTURA No. " & FA.Serie & "-"
       Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
       Ln_No = 1
       TxtAfiliado = ""
       TxtBeneficiario = ""
       TxtBeneficiario2 = ""
       TxtAfiliado.SetFocus
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
          TxtEfectivo = "0.00"
          TxtEfectivo.SetFocus
     Case vbKeyReturn
          TextCant.SetFocus
   End Select
  If KeyCode = vbKeyF1 Then
     With AdoArticulo.Recordset
      If .RecordCount Then
         .MoveFirst
         .Find ("Nom_Art = '" & DCArticulo & "' ")
          If Not .EOF Then MsgBox .fields("Producto") & ":" & vbCrLf & .fields("Ayuda")
      End If
     End With
  End If
End Sub

Private Sub DCArticulo_LostFocus()
  Codigos = Ninguno
  If Leer_Codigo_Inv(DCArticulo.Text, FechaSistema, Cod_Bodega) Then DatosArticulos
End Sub

Private Sub DCBodega_LostFocus()
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .fields("CodBod")
   End If
  End With
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
  Cod_Bodega = Ninguno
  FA.Fecha = MBFecha
  FechaValida MBFecha
  Grupo_Inv = Ninguno
  Ln_No = 1
  Cant_Item_PV = 30
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  CodigoCliente = ""
  NombreCliente = ""
  DireccionCli = " S/N"
  TextBenef = NombreCliente
  FacturaFarmacia.Caption = FacturaFarmacia.Caption & " (" & TipoFactura & ")"
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  TextCant.Text = "0"
  TextVUnit.Text = "0"
  LabelVTotal.Caption = "0"
  Modificar = False
  Bandera = True
  Mifecha = BuscarFecha(FechaSistema)
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  SerieFactura = "999999"
  Autorizacion = "9999999999"
  FA.Fecha = MBFecha
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact = '" & TipoFactura & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
       & "ORDER BY Codigo "
  Select_Adodc AdoLinea, sSQL
  With AdoLinea.Recordset
    'MsgBox sSQL & vbCrLf & .RecordCount
    If .RecordCount > 0 Then
        FA.Cod_CxC = .fields("Codigo")
        Lineas_De_CxC FA
    Else
        MsgBox "Falta Organizar la CxC en Puntos de Venta." & vbCrLf _
             & "Salga de este proceso y llame al su técnico" & vbCrLf _
             & "o al Contador de su Organizacion."
    End If
  End With
  FacturaFarmacia.Caption = "INGRESAR FACTURA"
  Label1.Caption = FA.Autorizacion & " - FACTURA No. " & FA.Serie & "-"
  Label3.Caption = "I.V.A. " & Format(Porc_IVA * 100, "#0.00") & "%"
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  
  TextFacturaNo = FA.Factura

  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SQLDec = "PRECIO 4|CORTE 5|."
  Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
    
  MBFecha = FechaSistema
  sSQL = "SELECT * " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodBod "
  SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega Like '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .fields("CodBod")
   End If
  End With
  
  sSQL = "SELECT Producto & ' -> ' & Codigo_Inv As Nom_Art " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'I' "
  If Len(Grupo_Inv) > 1 Then sSQL = sSQL & "AND MID(Codigo_Inv,1,2) = '" & Grupo_Inv & "' "
  sSQL = sSQL & "ORDER BY Producto,Codigo_Inv "
  SelectDB_List DLGrupo, AdoGrupo, sSQL, "Nom_Art"
  
  sSQL = "SELECT Producto, Codigo_Inv, Codigo_Barra " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' "
  If Len(Grupo_Inv) > 1 Then sSQL = sSQL & "AND MID(Codigo_Inv,1,2) = '" & Grupo_Inv & "' "
  If TipoFactura = "CP" Then
     sSQL = sSQL & "AND Cta_Inventario = '0' "
  Else
     sSQL = sSQL & "AND LEN(Cta_Inventario) >= 1 "
  End If
  sSQL = sSQL & "ORDER BY Producto,Codigo_Inv "
  SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
  FacturaFarmacia.WindowState = 2
  RatonNormal
  If AdoArticulo.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Productos de Venta"
     Unload FacturaFarmacia
  End If
End Sub

Private Sub Form_Deactivate()
  FacturaFarmacia.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoBenef
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFactura
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
  
  SRI_Obtener_Datos_Comprobantes_Electronicos
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  FechaTexto1 = MBFecha
  FA.Fecha = MBFecha
  FA.Fecha_V = MBFecha
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
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
      Real1 = Redondear(CCur(TextCant) * CCur(TextVUnit), 2)
      With AdoAsientoF.Recordset
        If Real1 > 0 Then
           Select Case TipoFactura
             Case "NV", "PV": Real3 = 0
             Case Else
                  If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
           End Select
           If Len(TxtDocumentos) > 1 Then Producto = Producto & " - " & TxtDocumentos
           SetAddNew AdoAsientoF
           SetFields AdoAsientoF, "CODIGO", Codigos
           SetFields AdoAsientoF, "CODIGO_L", CodigoL
           SetFields AdoAsientoF, "PRODUCTO", Mid(Producto, 1, 45)
           SetFields AdoAsientoF, "CANT", CCur(TextCant)
           SetFields AdoAsientoF, "PRECIO", CCur(TextVUnit)
           SetFields AdoAsientoF, "PRECIO2", CCur(DatInv.PVP2)
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

Public Sub ProcGrabar()
    DGAsientoF.Visible = False
    Codigo1 = TxtAfiliado
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
       TxtAfiliado = NombreCliente
    End If
   'Insertamos el beneficiario completo del CENTRO MEDICO
    sSQL = "SELECT * " _
         & "FROM Clientes_Matriculas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '" & FA.CodigoC & "' "
    Select_Adodc AdoFactura, sSQL
    If AdoFactura.Recordset.RecordCount <= 0 Then
       SetAdoAddNew "Clientes_Matriculas"
       SetAdoFields "T", Normal
       SetAdoFields "Codigo", FA.CodigoC
       SetAdoFields "TD", "R"
       SetAdoFields "Cedula_R", "1790764575001"
       SetAdoFields "Representante", "CENTRO MEDICO MATERNAL PAEZ ALMEIDA Y NARANJO"
       SetAdoFields "Lugar_Trabajo_R", "GARCIA MORENO Y ESMERALDAS"
       SetAdoFields "Telefono_R", "022282950"
       SetAdoFields "Telefono_RS", "022282950"
       SetAdoFields "Grupo_No", Grupo_No
       SetAdoFields "Item", NumEmpresa
       SetAdoFields "Periodo", Periodo_Contable
       SetAdoFields "CodigoU", CodigoUsuario
       SetAdoUpdate
       TBeneficiario = Leer_Datos_Cliente_SP(FA.CodigoC)
    End If
    
'''    sSQL = "UPDATE Clientes " _
'''         & "SET DireccionT = 'GARCIA MORENO Y ESMERALDAS' " _
'''         & "WHERE Codigo = '" & FA.CodigoC & "' "
'''    Ejecutar_SQL_SP sSQL

 'Seteamos los encabezados para las facturas
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
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
     FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
     FA.Tipo_PRN = "FM"
     Grabar_Factura FA, True
     
    'Abono de Factura
     TA.T = Normal
     TA.TP = FA.TC
     TA.Serie = FA.Serie
     TA.Autorizacion = FA.Autorizacion
     TA.CodigoC = FA.CodigoC
     TA.Factura = FA.Factura
     TA.Fecha = FA.Fecha
     TA.Cta = Cta_CajaG
     TA.Cta_CxP = FA.Cta_CxP
     TA.Banco = "ABONO POR VERIFICAR"
     TA.Cheque = UCase$(Grupo_No)
     TA.Abono = FA.Total_MN
     Grabar_Abonos TA
    
    'Autorizamos la factura i/o Guia de Remision
     If Len(FA.Autorizacion) = 13 Then SRI_Crear_Clave_Acceso_Facturas FA, True, , True
     RatonNormal
     Ln_No = 1
     Imprimir_Facturas_CxC FacturaFarmacia, FA
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
       Precio = Redondear(DatInv.PVP, 2)
       DatInv.PVP2 = Redondear(DatInv.PVP2, 2)
       LabelStock.Caption = DatInv.Stock
       BanIVA = DatInv.IVA
       TextVUnit = Format(Precio, "#,##0.0000")
       If TipoFactura = "NV" Then BanIVA = False
       DCArticulo.Text = Producto
   End If
  End With
End Sub

Private Sub TxtAfiliado_GotFocus()
   MarcarTexto TxtAfiliado
End Sub

Private Sub TxtAfiliado_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtAfiliado_LostFocus()
Dim EncontroCliente As Boolean
  EncontroCliente = False
  Codigo1 = TxtAfiliado
  TBeneficiario.Patron_Busqueda = Codigo1
  TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
  'MsgBox TBeneficiario.Codigo
  '0705280113
  If TBeneficiario.Codigo = Ninguno Then
     TBeneficiario.Patron_Busqueda = Codigo1
     FBuscarClientes.Show 1
     If TBeneficiario.Codigo = Ninguno Then
        NombreCliente = Codigo1
        Nuevo = True
        FClientesFlash.Show 1
        TxtAfiliado = NombreCliente
        TxtAfiliado.SetFocus
     Else
        EncontroCliente = True
     End If
  Else
     EncontroCliente = True
  End If
  If EncontroCliente Then
     CodigoCliente = TBeneficiario.Codigo
     NombreCliente = TBeneficiario.Cliente
     CICliente = TBeneficiario.CI_RUC
     Grupo_No = TBeneficiario.Grupo_No
     TipoDoc = TBeneficiario.TD
     FA.CodigoC = CodigoCliente
     FA.Cliente = NombreCliente
     FA.CI_RUC = CICliente
     CmdCIBenef1.Caption = TBeneficiario.CI_RUC
     TxtAfiliado = NombreCliente
  End If
End Sub

Private Sub TxtBeneficiario_GotFocus()
   TxtBeneficiario = TxtAfiliado
   MarcarTexto TxtBeneficiario
End Sub

Private Sub TxtBeneficiario_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtBeneficiario_LostFocus()
Dim EncontroCliente As Boolean
  EncontroCliente = False
  Codigo2 = TxtBeneficiario
  If Codigo2 = "" Then Codigo2 = Ninguno
  TBeneficiario.Patron_Busqueda = Codigo2
  TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
  If TBeneficiario.Codigo = Ninguno Then
     TBeneficiario.Patron_Busqueda = Codigo2
     FBuscarClientes.Show 1
     If TBeneficiario.Codigo = Ninguno Then
        NombreCliente = Codigo2
        Nuevo = True
        FClientesFlash.Show 1
        TxtBeneficiario = NombreCliente
        TxtBeneficiario.SetFocus
     Else
        EncontroCliente = True
     End If
  Else
     EncontroCliente = True
  End If
  If EncontroCliente Then
     CodigoBenef = TBeneficiario.Codigo
     FA.CodigoB = CodigoBenef
     CmdCIBenef2.Caption = TBeneficiario.CI_RUC
     TxtBeneficiario = TBeneficiario.Cliente
  End If
End Sub

Private Sub TxtBeneficiario2_GotFocus()
   MarcarTexto TxtBeneficiario2
End Sub

Private Sub TxtBeneficiario2_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtBeneficiario2_LostFocus()
Dim EncontroCliente As Boolean
  EncontroCliente = False
  Codigo3 = TxtBeneficiario2
  If Codigo3 = "" Then Codigo3 = Ninguno
  TBeneficiario.Patron_Busqueda = Codigo3
  TBeneficiario = Leer_Datos_Cliente_SP(TBeneficiario.Patron_Busqueda)
  If TBeneficiario.Codigo = Ninguno Then
     TBeneficiario.Patron_Busqueda = Codigo3
     FBuscarClientes.Show 1
     If TBeneficiario.Codigo = Ninguno Then
        NombreCliente = Codigo3
        Nuevo = True
        FClientesFlash.Show 1
        TBeneficiario.CI_RUC = Ninguno
        EncontroCliente = True
     Else
        EncontroCliente = True
     End If
  Else
     EncontroCliente = True
  End If
  If EncontroCliente Then
     CodigoBenef2 = TBeneficiario.Codigo
     FA.CodigoA = CodigoBenef2
     CmdCIBenef3.Caption = TBeneficiario.CI_RUC
     TxtBeneficiario2 = TBeneficiario.Cliente
  End If
End Sub

Private Sub TxtDocumentos_GotFocus()
  MarcarTexto TxtDocumentos
End Sub

Private Sub TxtDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocumentos_LostFocus()
  TextoValido TxtDocumentos
End Sub

Private Sub TxtEfectivo_GotFocus()
  MarcarTexto TxtEfectivo
End Sub

Private Sub TxtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_Change()
'  If Val(TextCotiza) > 0 Then
'     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_FacturaME, "#,##0.00")
'  Else
'     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_Factura, "#,##0.00")
'  End If
End Sub

Private Sub TxtEfectivo_LostFocus()
  TextoValido TxtEfectivo, True, , 2
  If Val(TextCotiza) > 0 Then
     LblCambio.Caption = Format(Val(CCur(TxtEfectivo)) - Total_FacturaME, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - Total_FacturaME) >= 0 Then Command3.SetFocus
  Else
     If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then Command3.SetFocus
  End If
End Sub
