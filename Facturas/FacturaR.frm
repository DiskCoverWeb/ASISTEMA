VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturaReembolso 
   BackColor       =   &H00C0C0C0&
   Caption         =   "2023-09-30 00:00:00.000"
   ClientHeight    =   11970
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   16305
   WindowState     =   1  'Minimized
   Begin VB.TextBox TxtAutProvee 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   9345
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   2100
      Width           =   6840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&+"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6615
      Picture         =   "FacturaR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1785
      Width           =   540
   End
   Begin VB.TextBox TxtFacturaP 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8085
      MaxLength       =   9
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "FacturaR.frx":5C12
      Top             =   2100
      Width           =   1170
   End
   Begin VB.TextBox TxtSerieP 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7245
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "FacturaR.frx":5C1E
      Top             =   2100
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1275
      Left            =   105
      TabIndex        =   33
      Top             =   8925
      Width           =   16080
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
         Height          =   855
         Left            =   14280
         Picture         =   "FacturaR.frx":5C27
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   210
         Width           =   1590
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
         Height          =   855
         Left            =   12495
         Picture         =   "FacturaR.frx":64F1
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   210
         Width           =   1590
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
         TabIndex        =   34
         Top             =   210
         Width           =   2430
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
         Left            =   2940
         TabIndex        =   36
         Top             =   210
         Width           =   2430
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
         Left            =   105
         TabIndex        =   35
         Top             =   630
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
         Left            =   2940
         TabIndex        =   37
         Top             =   630
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
         Left            =   5775
         TabIndex        =   38
         Top             =   210
         Width           =   2430
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
         Left            =   8610
         TabIndex        =   40
         Top             =   210
         Width           =   2430
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
         Left            =   5775
         TabIndex        =   39
         Top             =   630
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
         Left            =   8610
         TabIndex        =   41
         Top             =   630
         Width           =   2430
      End
   End
   Begin VB.TextBox TxtObservacion 
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
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   8505
      Width           =   7995
   End
   Begin VB.TextBox TxtNota 
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
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   8505
      Width           =   7995
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7455
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "FacturaR.frx":67FB
      Top             =   2835
      Width           =   1170
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6405
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "FacturaR.frx":6800
      Top             =   2835
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   10290
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   2835
      Width           =   5895
   End
   Begin VB.TextBox TextFacturaNo 
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
      Height          =   390
      Left            =   14385
      TabIndex        =   8
      Text            =   "9999999999"
      Top             =   525
      Width           =   1800
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FacturaR.frx":6802
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   14
      Top             =   2100
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   16711680
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
      Bindings        =   "FacturaR.frx":681C
      Height          =   3585
      Left            =   105
      TabIndex        =   28
      Top             =   3255
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   6324
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Serif"
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
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
      Top             =   5250
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
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   2625
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2625
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FacturaR.frx":6836
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   1470
      TabIndex        =   3
      Top             =   420
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CONSUMIDOR FINAL"
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
   Begin MSDataListLib.DataCombo DCTipoPago 
      Bindings        =   "FacturaR.frx":684D
      DataSource      =   "AdoTipoPago"
      Height          =   360
      Left            =   9345
      TabIndex        =   12
      Top             =   1365
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   525
      Top             =   5565
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "TipoPago"
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
   Begin MSDataListLib.DataCombo DCTipoComprobante 
      Bindings        =   "FacturaR.frx":6867
      DataSource      =   "AdoTipoComp"
      Height          =   315
      Left            =   105
      TabIndex        =   46
      ToolTipText     =   $"FacturaR.frx":6888
      Top             =   2835
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AUTORIZACON DEL COMPROBANTE"
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
      TabIndex        =   47
      Top             =   2520
      Width           =   6210
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AUTORIZACON DEL COMPROBANTE"
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
      Left            =   9345
      TabIndex        =   45
      Top             =   1785
      Width           =   6840
   End
   Begin VB.Label Label34 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE PAGO"
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
      Left            =   9345
      TabIndex        =   11
      Top             =   1050
      Width           =   6840
   End
   Begin VB.Label LblDireccion 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   1365
      Width           =   9150
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   8085
      TabIndex        =   18
      Top             =   1785
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      Left            =   7245
      TabIndex        =   16
      Top             =   1785
      Width           =   750
   End
   Begin VB.Label Label29 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION"
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
      TabIndex        =   9
      Top             =   1050
      Width           =   9150
   End
   Begin VB.Label Label24 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OBSERVACION:"
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
      Left            =   8190
      TabIndex        =   31
      Top             =   8190
      Width           =   7995
   End
   Begin VB.Label Label21 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOTA:"
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
      TabIndex        =   29
      Top             =   8190
      Width           =   7995
   End
   Begin VB.Label LblRUC 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   11130
      TabIndex        =   5
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CI/RUC/PAS."
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
      Left            =   11130
      TabIndex        =   4
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label LblSerie 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999-999"
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
      Left            =   13125
      TabIndex        =   7
      Top             =   525
      Width           =   1275
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8715
      TabIndex        =   25
      Top             =   2835
      Width           =   1485
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUBTOTAL"
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
      TabIndex        =   24
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.V.P"
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
      TabIndex        =   22
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   6405
      TabIndex        =   20
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DETALLE DEL REEMBOLSO"
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
      TabIndex        =   26
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
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
      Width           =   9570
   End
   Begin VB.Label LabelStockArt 
      BackColor       =   &H00C0FFFF&
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
      Left            =   105
      TabIndex        =   13
      Top             =   1785
      Width           =   6420
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   13125
      TabIndex        =   6
      Top             =   105
      Width           =   3060
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
End
Attribute VB_Name = "FacturaReembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Grupo_Inv As String

Dim Total_PV As Currency
Dim SaldoPendiente As Currency

Dim CantSaldoPendiente As Integer

Dim ParpadearSaldo As Boolean
Dim SiTPCJ As Boolean
Dim SiTPBA As Boolean

Private Sub Command2_Click()
    NombreCliente = "NUEVO PROVEEDOR"
    Nuevo = True
    FClientesFlash.Show 1
End Sub

Private Sub DCArticulo_GotFocus()
    Calculos_Totales_Factura FA
    LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
    LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
    LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
    LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    Select Case KeyCode
      Case vbKeyEscape
           Calculos_Totales_Factura FA
           LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
           LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
           LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
           LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
           Command3.SetFocus
      Case vbKeyReturn
           SiguienteControl
    End Select
End Sub

Private Sub DCArticulo_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCArticulo.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Cliente, Codigo, CI_RUC " _
            & "FROM Clientes " _
            & "WHERE TD = 'R' "
       If IsNumeric(Busqueda) Then
          sSQL = sSQL & "AND CI_RUC LIKE '" & Busqueda & "%' "
       Else
          sSQL = sSQL & "AND Cliente LIKE '%" & Busqueda & "%' "
       End If
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoArticulo, sSQL
    End If
End Sub

Private Sub DCArticulo_LostFocus()
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       CodigoC = "CONSUMIDOR FINAL"
       CodigoB = "9999999999"
       CodigoP = "9999999999999"
       If IsNumeric(DCArticulo.Text) Then .Find ("CI_RUC = '" & DCArticulo.Text & "'") Else .Find ("Cliente = '" & DCArticulo.Text & "'")
       If Not .EOF Then
          CodigoC = .fields("Cliente")
          CodigoB = .fields("Codigo")
          CodigoP = .fields("CI_RUC")
       End If
   End If
  End With
End Sub

Private Sub DCCliente_GotFocus()
   FA.DireccionS = Ninguno
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Cliente,Codigo,CI_RUC,TD,Grupo,Direccion,DirNumero " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then
          If Len(Busqueda) = 4 Then sSQL = sSQL & "WHERE DirNumero = '" & Busqueda & "' " Else sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' "
       Else
          sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       End If
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoBenef, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
Dim DireccionC As String
  SaldoPendiente = 0
  CantSaldoPendiente = 0
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       If IsNumeric(DCCliente.Text) Then
         .Find ("CI_RUC = '" & DCCliente.Text & "'")
       Else
         .Find ("Cliente = '" & DCCliente.Text & "'")
       End If
       If Not .EOF Then
          CodigoBenef = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          DireccionC = .fields("Direccion")
          DCCliente.Text = .fields("Cliente")
          LblRUC.Caption = .fields("CI_RUC")
          LblDireccion.Caption = DireccionC
                    
          sSQL = "SELECT COUNT(Factura) CantFact, SUM(Saldo_MN) As TSaldo_MN " _
               & "FROM Facturas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND CodigoC = '" & CodigoCliente & "' "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             If Not IsNull(AdoAux.Recordset.fields("TSaldo_MN")) Then
                SaldoPendiente = AdoAux.Recordset.fields("TSaldo_MN")
                CantSaldoPendiente = AdoAux.Recordset.fields("CantFact")
             End If
          End If
       Else
          NombreCliente = DCCliente.Text
          FacturaReembolso.Visible = False
          Nuevo = True
          MsgBox "Cliente No existe"
          FClientesFlash.Show 1
          FacturaReembolso.Visible = True
          Listar_Clientes_FR
       End If
   Else
       NombreCliente = DCCliente.Text
       Nuevo = True
       FClientesFlash.Show 1
       Listar_Clientes_FR
   End If
  End With
End Sub

Private Sub DCTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
 'Validar_Porc_IVA MBFecha
  FechaTexto1 = MBFecha
  LblSerie.Caption = SerieFactura & "-"
  NumComp = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  TextFacturaNo.Text = Format$(NumComp, "000000000")
End Sub

Public Sub Listar_Clientes_FR()
  sSQL = "SELECT TOP 50 Cliente,Codigo,CI_RUC,TD,Grupo,Direccion " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND FA <> " & Val(adFalse) & " " _
       & "UNION " _
       & "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,Direccion " _
       & "FROM Clientes " _
       & "WHERE Codigo = '9999999999' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
End Sub

Private Sub Command1_Click()
  Unload FacturaReembolso
End Sub

Private Sub Command3_Click()
    FechaValida MBFecha
    FechaTexto = MBFecha
    Calculos_Totales_Factura FA

    Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
             & "Comprobante (" & FA.TC & ")  No. " & TextFacturaNo.Text
    Titulo = "Formulario de Grabacion"
    If BoxMensaje = vbYes Then
       FA.Nota = TxtNota
       FA.Observacion = TxtObservacion
       FA.Tipo_Pago = SinEspaciosIzq(DCTipoPago)
       FA.TDT = 41
       Moneda_US = False
       TextoFormaPago = "PENDIENTE"
       ProcGrabar
       
       LblSerie.Caption = SerieFactura & "-"
       NumComp = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
       TextFacturaNo.Text = Format$(NumComp, "000000000")
    
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
       Ln_No = 1
       MBFecha.SetFocus
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
 'Facturas de Reembolso
 '---------------------
  TipoFactura = "FR"
 '---------------------
  Grupo_Inv = Ninguno
  Ln_No = 1
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  CodigoCliente = "9999999999"
  NombreCliente = "CONSUMIDOR FINAL"
  DireccionCli = " S/N"
  DCCliente.Text = NombreCliente
  FacturaReembolso.Caption = "FACTURA POR REEMBOLSO DE GASTOS (" & TipoFactura & ")"
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  TextCant.Text = "0"
  TextVUnit.Text = "0"
  LabelVTotal.Caption = "0"
  Modificar = False
  Bandera = True
  Mifecha = BuscarFecha(FechaSistema)
  FA.TC = TipoFactura
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  Autorizacion = "9999999999"
  sSQL = "SELECT CxC, Codigo, Autorizacion, Fact, Serie " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact = 'FR' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "ORDER BY Serie, Codigo "
  Select_Adodc AdoLinea, sSQL
  With AdoLinea.Recordset
   If .RecordCount > 0 Then
        Cta_Cobrar = .fields("CxC")
        CodigoL = .fields("Codigo")
        FA.TC = .fields("Fact")
        FA.Serie = .fields("Serie")
        FA.Autorizacion = .fields("Autorizacion")
        FA.Cta_CxP = Cta_Cobrar
        FA.Cod_CxC = CodigoL
        Autorizacion = FA.Autorizacion
        SerieFactura = FA.Serie
        
        Label1.Caption = " FACTURA No."
        Label3.Caption = " I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"

        LblSerie.Caption = SerieFactura & "-"
        NumComp = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
        TextFacturaNo.Text = Format$(NumComp, "000000000")
        
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        SQLDec = "PRECIO 4|CORTE 5|."
        Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
        RatonNormal
        Listar_Clientes_FR
        MBFecha.Text = FechaSistema
        Cod_Bodega = "01"
        
        sSQL = "SELECT TOP 50 Cliente,Codigo,CI_RUC,TD,Grupo,Direccion,DirNumero " _
             & "FROM Clientes " _
             & "WHERE TD = 'R' " _
             & "ORDER BY Cliente "
        SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Cliente"
        
        sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
             & "FROM Tabla_Referenciales_SRI " _
             & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
             & "AND Codigo IN ('01','16','17','18','19','20','21') " _
             & "ORDER BY Codigo "
        SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
        
        sSQL = "SELECT * " _
             & "FROM Tipo_Comprobante " _
             & "WHERE Tipo_Comprobante_Codigo IN (" & Cadena & ") " _
             & "AND TC = 'TDC' "
        If TipoBenef = "R" Then
           sSQL = sSQL & "AND R <> " & Val(adFalse) & " "
        Else
           sSQL = sSQL & "AND C <> " & Val(adFalse) & " "
        End If
        sSQL = sSQL & "ORDER BY Tipo_Comprobante_Codigo "
        SelectDB_Combo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion", , "FCimprasAT"
        

        SaldoPendiente = 0
        CantSaldoPendiente = 0
        ParpadearSaldo = True
        FacturaReembolso.WindowState = 2
        If AdoArticulo.Recordset.RecordCount <= 0 Then
           MsgBox "No existen RUC asignados"
           Unload FacturaReembolso
        End If
        Encontro = Leer_Codigo_Inv("99.41", FechaSistema, Cod_Bodega)
        If Encontro Then
           DatosArticulos
        Else
           MsgBox "Falta asignar codigos de Facturacion de Reembolsos de Gastos"
           Unload FacturaReembolso
        End If
   Else
        MsgBox "Falta Organizar la CxC en Reembolsos de Gastos." & vbCrLf _
             & "Salga de este proceso y llame al su técnico" & vbCrLf _
             & "o al Contador de su Organizacion."
        Unload FacturaReembolso
   End If
  End With
End Sub

Private Sub Form_Deactivate()
  FacturaReembolso.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoBenef
  ConectarAdodc AdoLinea
  ConectarAdodc AdoFactura
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
  ConectarAdodc AdoTipoPago
  Encerar_Factura FA
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
  Codigos = Ninguno
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
End Sub

Private Sub TextCant_LostFocus()
  Cantidad = Val(TextCant)
End Sub

Private Sub TextVUnit_Change()
   If IsNumeric(TextVUnit) And IsNumeric(TextCant) Then
      If Val(TextVUnit) = 0 Then TextVUnit = "0.01"
      Real1 = CCur(TextCant) * CCur(TextVUnit)
      LabelVTotal.Caption = Format$(Real1, "#,##0.00")
   Else
      LabelVTotal.Caption = "0.0000"
   End If
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
   TextoValido TextVUnit, True, , 2
   Mensajes = "Ingreso de Reembolso Con IVA?" & vbCrLf
   Titulo = "FORMULARIO DE CONFIRMACION"
   If BoxMensaje = vbYes Then BanIVA = True Else BanIVA = False
   TxtDocumentos.SetFocus
End Sub

Public Sub ProcGrabar()
 DGAsientoF.Visible = False
 FA.Porc_IVA = Porc_IVA
 FA.Gavetas = 0
 'Seteamos los encabezados para las facturas
  Calculos_Totales_Factura FA
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     FechaTexto = MBFecha
     FA.Fecha = FechaTexto
     FA.CodigoC = CodigoCliente
     HoraTexto = Format$(Time, FormatoTimes)
     Total_FacturaME = 0
     Moneda_US = False
     Total_Factura = Redondear(FA.Sin_IVA + FA.Con_IVA - FA.Descuento - FA.Descuento2 + FA.Total_IVA + FA.Servicio, 2)
     Total_FacturaME = Total_Factura
     If Moneda_US Then Total_Factura = Redondear(Total_Factura * Dolar, 2) Else Total_FacturaME = 0
     Saldo = Total_Factura
     Saldo_ME = Total_FacturaME
     If Saldo < 0 Then Saldo = 0
     FA.Nuevo_Doc = True
     FA.EsPorReembolso = True
     Factura_No = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
     FA.Factura = Factura_No
     TipoFactura = FA.TC
     
     sSQL = "DELETE * " _
          & "FROM Detalle_Factura " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "DELETE * " _
          & "FROM Facturas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     Ejecutar_SQL_SP sSQL
     
     TextoFormaPago = PagoCred
     T = Pendiente
    'Grabamos el numero de factura
     RatonNormal
     Grabar_Factura FA, True

    'Actualizamos el saldo de la factura
     Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
     
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
     If Len(FA.Autorizacion) >= 13 Then
        SRI_Crear_Clave_Acceso_Facturas FA, True
     Else
        MsgBox "No es factura de Reembolso"
     End If
     Listar_Clientes_FR
     NombreCliente = "CONSUMIDOR FINAL"
     DCCliente.Text = NombreCliente
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub DatosArticulos()
''  With AdoArticulo.Recordset
''   If .RecordCount > 0 Then
    Codigos = DatInv.Codigo_Inv
    Producto = DatInv.Producto
    Cta_Ventas = DatInv.Cta_Ventas
    Precio = DatInv.PVP
    BanIVA = DatInv.IVA
    TextVUnit = Format$(Precio, "#,##0.0000")
    If TipoFactura = "NV" Then BanIVA = False
    LabelStockArt.Caption = "PRODUCTO" & String(93 - Len(DatInv.Codigo_Inv), " ") & DatInv.Codigo_Inv
    'DCArticulo.Text = Producto
''   End If
''  End With
End Sub

Private Sub TxtAutProvee_GotFocus()
  MarcarTexto TxtAutProvee
End Sub

Private Sub TxtAutProvee_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAutProvee_LostFocus()
  TextoValido TxtAutProvee, True, , 0
End Sub

Private Sub TxtDocumentos_GotFocus()
  MarcarTexto TxtDocumentos
End Sub

Private Sub TxtDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub TxtDocumentos_LostFocus()
Dim Grabar_PV As Boolean
Dim ProductoAux As String
   TextoValido TxtDocumentos
   DatInv.Codigo_Inv = "99.41"
      LabelVTotal.Caption = Format$(Real1, "#,##0.00")
      Real1 = 0: Real2 = 0: Real3 = 0
      If IsNumeric(TextVUnit) And IsNumeric(TextCant) Then Real1 = CCur(TextCant) * CCur(TextVUnit)
      LabelVTotal.Caption = Format$(Real1, "#,##0.00")
      If Real1 > 0 Then
         If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
         SetAddNew AdoAsientoF
         SetFields AdoAsientoF, "CODIGO", DatInv.Codigo_Inv
         SetFields AdoAsientoF, "CODIGO_L", CodigoL
         SetFields AdoAsientoF, "PRODUCTO", CodigoC
         SetFields AdoAsientoF, "Tipo_Hab", MidStrg(TxtDocumentos, 1, 40)
         SetFields AdoAsientoF, "RUTA", CodigoP
         SetFields AdoAsientoF, "Serie_No", TxtSerieP & "-" & TxtFacturaP
         SetFields AdoAsientoF, "Codigo_B", CodigoB
         SetFields AdoAsientoF, "CANT", CCur(TextCant)
         SetFields AdoAsientoF, "PRECIO", CCur(TextVUnit)
         SetFields AdoAsientoF, "TOTAL", Real1
         SetFields AdoAsientoF, "Total_IVA", Real3
         SetFields AdoAsientoF, "Item", NumEmpresa
         SetFields AdoAsientoF, "CodigoU", CodigoUsuario
         SetFields AdoAsientoF, "CodBod", Cod_Bodega
         SetFields AdoAsientoF, "A_No", Ln_No
         SetUpdate AdoAsientoF
         Ln_No = Ln_No + 1
      End If
   TextCant.Text = "0"
   DCArticulo.SetFocus
End Sub

Private Sub TxtFacturaP_GotFocus()
  MarcarTexto TxtFacturaP
End Sub

Private Sub TxtFacturaP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFacturaP_LostFocus()
  TxtFacturaP = Format(Val(TxtFacturaP), "000000000")
End Sub

Private Sub TxtNota_GotFocus()
  MarcarTexto TxtNota
End Sub

Private Sub TxtNota_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNota_LostFocus()
  TextoValido TxtNota, , True
End Sub

Private Sub TxtObservacion_GotFocus()
   MarcarTexto TxtObservacion
End Sub

Private Sub TxtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtObservacion_LostFocus()
  TextoValido TxtObservacion, , True
End Sub

Private Sub TxtSerieP_GotFocus()
  MarcarTexto TxtSerieP
End Sub

Private Sub TxtSerieP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSerieP_LostFocus()
  TextoValido TxtSerieP
  If Len(TxtSerieP) < 6 Then TxtSerieP = String(6 - Len(TxtSerieP), "0") & TxtSerieP
End Sub
