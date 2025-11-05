VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FFletes 
   Caption         =   "INGRESOS DE FLETES"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Command1"
      Height          =   960
      Left            =   10185
      TabIndex        =   35
      Top             =   2100
      Width           =   1065
   End
   Begin VB.TextBox TxtKmF 
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
      Left            =   9345
      TabIndex        =   27
      Text            =   "0"
      Top             =   1365
      Width           =   1905
   End
   Begin VB.TextBox TxtKmI 
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
      Left            =   9345
      TabIndex        =   25
      Text            =   "0"
      Top             =   1050
      Width           =   1905
   End
   Begin VB.TextBox TxtFlete 
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
      Left            =   9345
      TabIndex        =   23
      Text            =   "0"
      Top             =   735
      Width           =   1905
   End
   Begin VB.TextBox TxtPlaca 
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
      Left            =   8295
      MaxLength       =   20
      TabIndex        =   21
      Top             =   420
      Width           =   2955
   End
   Begin VB.TextBox TxtCond 
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
      Left            =   5355
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1680
      Width           =   1380
   End
   Begin VB.TextBox TxtCarga 
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
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1680
      Width           =   2640
   End
   Begin MSDataListLib.DataCombo DCAyudante 
      Bindings        =   "FFletes.frx":0000
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1470
      TabIndex        =   15
      ToolTipText     =   "<Ctrl+F> Lista los Fletes de este Cliente"
      Top             =   1365
      Width           =   5265
      _ExtentX        =   9287
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
   Begin MSDataListLib.DataCombo DCConductor 
      Bindings        =   "FFletes.frx":001A
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1470
      TabIndex        =   13
      ToolTipText     =   "<Ctrl+F> Lista los Fletes de este Cliente"
      Top             =   1050
      Width           =   5265
      _ExtentX        =   9287
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
   Begin MSDataListLib.DataCombo DCDestino 
      Bindings        =   "FFletes.frx":0034
      DataSource      =   "AdoDestino"
      Height          =   315
      Left            =   1470
      TabIndex        =   11
      Top             =   735
      Width           =   5265
      _ExtentX        =   9287
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FFletes.frx":004D
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1470
      TabIndex        =   9
      ToolTipText     =   "<Ctrl+F> Lista los Fletes de este Cliente"
      Top             =   420
      Width           =   5265
      _ExtentX        =   9287
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
   Begin MSDataListLib.DataCombo DCFacturaNo 
      Bindings        =   "FFletes.frx":0067
      DataSource      =   "AdoNumero"
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
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
   Begin MSAdodcLib.Adodc AdoFletes 
      Height          =   330
      Left            =   105
      Top             =   6405
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Fletes"
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
   Begin MSDataGridLib.DataGrid DGFletes 
      Bindings        =   "FFletes.frx":007F
      Height          =   4320
      Left            =   105
      TabIndex        =   34
      Top             =   2100
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   7620
      _Version        =   393216
      AllowUpdate     =   0   'False
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   4410
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   7035
      TabIndex        =   5
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   525
      Top             =   3360
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoDestino 
      Height          =   330
      Left            =   525
      Top             =   3675
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
      Caption         =   "Destino"
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
      Left            =   525
      Top             =   3990
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
   Begin MSAdodcLib.Adodc AdoNumero 
      Height          =   330
      Left            =   525
      Top             =   4305
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
      Caption         =   "Numero"
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
   Begin VB.Label Label16 
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
      Height          =   330
      Left            =   9345
      TabIndex        =   29
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KILOMETRAJE RECORRIDO"
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
      Left            =   6720
      TabIndex        =   28
      Top             =   1680
      Width           =   2640
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KILOMETRAJE FINAL"
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
      Left            =   6720
      TabIndex        =   26
      Top             =   1365
      Width           =   2640
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KILOMETRAJE INICIAL"
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
      Left            =   6720
      TabIndex        =   24
      Top             =   1050
      Width           =   2640
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FLETE USD"
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
      Left            =   6720
      TabIndex        =   22
      Top             =   735
      Width           =   2640
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RUTA"
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
      Left            =   6720
      TabIndex        =   20
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   9345
      TabIndex        =   7
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FACTURA"
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
      Left            =   8295
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REFERENC."
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
      Left            =   4095
      TabIndex        =   18
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CARGA"
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
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " AYUDANTE"
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
      Top             =   1365
      Width           =   1380
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONDUCTOR"
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
      Top             =   1050
      Width           =   1380
   End
   Begin VB.Label Label19 
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
      Height          =   330
      Left            =   8085
      TabIndex        =   30
      Top             =   6405
      Width           =   2010
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL FACTURADO"
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
      TabIndex        =   31
      Top             =   6405
      Width           =   2010
   End
   Begin VB.Label Label9 
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
      Height          =   330
      Left            =   4095
      TabIndex        =   32
      Top             =   6405
      Width           =   2010
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PENDIENTES"
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
      Left            =   2730
      TabIndex        =   33
      Top             =   6405
      Width           =   1380
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA INICIAL"
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
      Width           =   1590
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA FINAL"
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
      TabIndex        =   4
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PLACA No."
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
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FLETE No."
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
      TabIndex        =   0
      Top             =   105
      Width           =   1380
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
      TabIndex        =   8
      Top             =   420
      Width           =   1380
   End
End
Attribute VB_Name = "FFletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
  Factura_No = Val(DCFacturaNo)
  Mensajes = "Grabar el Flete No. " & Factura_No
  Titulo = "PREGUNTA DE GRABACION"
  If BoxMensaje = vbYes Then
     RatonReloj
     sSQL = "DELETE * " _
          & "FROM Trans_Fletes " _
          & "WHERE Numero = " & Val(DCFacturaNo) & " "
     Ejecutar_SQL_SP sSQL
     SetAdoAddNew "Trans_Fletes"
     SetAdoFields "T", Normal
     SetAdoFields "Fecha_I", MBFechaI
     SetAdoFields "Fecha_F", MBFechaF
     SetAdoFields "Codigo_Inv", CodigoInv
     SetAdoFields "Numero", Val(DCFacturaNo)
     SetAdoFields "Carga", TxtCarga
     SetAdoFields "Ayudante", CodigoA
     SetAdoFields "Conductor", CodigoL
     SetAdoFields "Referencia", TxtCond
     SetAdoFields "Flete", Val(TxtFlete)
     SetAdoFields "Ruta", TxtPlaca
     SetAdoFields "Km_Inicial", Val(TxtKmI)
     SetAdoFields "Km_Final", Val(TxtKmF)
     SetAdoFields "CodigoC", CodigoCliente
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Periodo", Periodo_Contable
     SetAdoFields "Item", NumEmpresa
     SetAdoUpdate
     'MsgBox CodigoA & vbCrLf & CodigoL
     Listar_Fletes Val(DCFacturaNo)
     RatonNormal
  End If
End Sub

Private Sub CommandButton2_Click()
  Factura_No = Val(DCFacturaNo)
  Mensajes = "Anular el Flete No. " & Factura_No
  Titulo = "PREGUNTA DE ANULACION"
  If BoxMensaje = vbYes Then
     sSQL = "DELETE * " _
          & "FROM Trans_Fletes " _
          & "WHERE Numero = " & Val(DCFacturaNo) & " "
     Ejecutar_SQL_SP sSQL
  End If
End Sub

Private Sub CommandButton3_Click()
Dim SizeLetra As Integer

CEConLineas = ProcesarSeteos("HD")
SizeLetra = 10
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, AdoFletes, SizeLetra, TipoTimes, Orientacion_Pagina
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = True
With AdoFletes.Recordset
 If .RecordCount > 0 Then
''     Codigo1 = Ninguno
''     Codigo2 = Ninguno
''     If AdoClientes.Recordset.RecordCount Then
''        AdoClientes.Recordset.MoveFirst
''        AdoClientes.Recordset.Find ("Codigo = '" & .Fields("Ayudante") & "' ")
''        If Not AdoClientes.Recordset.EOF Then Codigo1 = AdoClientes.Recordset.Fields("Cliente")
''
''        AdoClientes.Recordset.MoveFirst
''        AdoClientes.Recordset.Find ("Codigo = '" & .Fields("Conductor") & "' ")
''        If Not AdoClientes.Recordset.EOF Then Codigo2 = AdoClientes.Recordset.Fields("Cliente")
''     End If
     Printer.FontSize = SetD(2).Porte
     PrinterTexto SetD(2).PosX, SetD(2).PosY, .fields("Fecha_I")
     Printer.FontSize = SetD(3).Porte
     PrinterTexto SetD(3).PosX, SetD(3).PosY, .fields("Fecha_F")
     Printer.FontSize = SetD(4).Porte
     PrinterTexto SetD(4).PosX, SetD(4).PosY, .fields("Cliente")
     Printer.FontSize = SetD(5).Porte
     PrinterTexto SetD(5).PosX, SetD(5).PosY, .fields("Nombre_Conductor")   'Codigo2
     Printer.FontSize = SetD(6).Porte
     PrinterTexto SetD(6).PosX, SetD(6).PosY, .fields("Placa")
     Printer.FontSize = SetD(7).Porte
     PrinterTexto SetD(7).PosX, SetD(7).PosY, .fields("Nombre_Ayudante")  ' Codigo1
     Printer.FontSize = SetD(8).Porte
     PrinterTexto SetD(8).PosX, SetD(8).PosY, .fields("Carga")
     Printer.FontSize = SetD(9).Porte
     PrinterTexto SetD(9).PosX, SetD(9).PosY, .fields("Ruta")
 End If
End With
RatonNormal
MensajeEncabData = ""
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Private Sub CommandButton4_Click()
  Unload FFletes
End Sub

Private Sub DCAyudante_LostFocus()
  CodigoA = Ninguno
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCAyudante & "' ")
       If Not .EOF Then CodigoA = .fields("Codigo")
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF Then
     With AdoClientes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente Like '" & DCCliente & "' ")
          If Not .EOF Then Codigos = .fields("Codigo")
      End If
     End With
     Listar_Fletes 0, Codigos
  End If
End Sub

Private Sub DCCliente_LostFocus()
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCCliente & "' ")
       If Not .EOF Then
          CodigoCliente = .fields("Codigo")
          NombreCliente = DCCliente
       Else
          MsgBox "No existen datos"
       End If
       
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCConductor_LostFocus()
  CodigoL = Ninguno
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCConductor & "' ")
       If Not .EOF Then CodigoL = .fields("Codigo")
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCDestino_LostFocus()
  With AdoDestino.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & DCDestino & "' ")
       If Not .EOF Then
          CodigoInv = .fields("Codigo_Inv")
       Else
          MsgBox "No existen datos"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFacturaNo_LostFocus()
  Factura_No = Val(DCFacturaNo)
  Listar_Fletes Factura_No
End Sub

Private Sub DGFletes_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyS Then
     Numero = DGFletes.Columns(2)
     Codigos = DGFletes.Columns(12)
     sSQL = "UPDATE Trans_Fletes " _
          & "SET Ok = " & Val(adTrue) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Numero = " & Numero & " "
     Ejecutar_SQL_SP sSQL
     Listar_Fletes 0, Codigos
  End If
  If CtrlDown And KeyCode = vbKeyN Then
     Numero = DGFletes.Columns(2)
     Codigos = DGFletes.Columns(12)
     sSQL = "UPDATE Trans_Fletes " _
          & "SET Ok <> " & Val(adTrue) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Numero = " & Numero & " "
     Ejecutar_SQL_SP sSQL
     Listar_Fletes 0, Codigos
  End If
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGFletes.Visible = False
     GenerarDataTexto FFletes, AdoFletes
     DGFletes.Visible = True
     DGFletes.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoClientes, sSQL, "Cliente"
  SelectDB_Combo DCAyudante, AdoClientes, sSQL, "Cliente"
  SelectDB_Combo DCConductor, AdoClientes, sSQL, "Cliente"
    
  sSQL = "SELECT * " _
       & "FROM Trans_Fletes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T = '" & Normal & "' " _
       & "ORDER BY Numero "
  SelectDB_Combo DCFacturaNo, AdoNumero, sSQL, "Numero"
  RatonNormal
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE TC = 'P' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "ORDER BY Producto,Codigo_Inv "
  SelectDB_Combo DCDestino, AdoDestino, sSQL, "Producto"
  Listar_Fletes Val(DCFacturaNo)
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoFletes
   ConectarAdodc AdoNumero
   ConectarAdodc AdoDestino
   ConectarAdodc AdoClientes
End Sub


Public Sub Listar_Fletes(NumeroFlete As Long, Optional CodigoC As String)
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  CodigoCliente = "."
  CodigoInv = "."
  TxtCarga = ""
  TxtCond = ""
  TxtFlete = "0"
  TxtPlaca = ""
  TxtKmI = "0"
  TxtKmF = "0"
  
  sSQL = "SELECT Cliente,Producto,TF.* " _
       & "FROM Clientes As C,Catalogo_Productos As CP,Trans_Fletes As TF " _
       & "WHERE TF.Item = '" & NumEmpresa & "' " _
       & "AND TF.Periodo = '" & Periodo_Contable & "' " _
       & "AND TF.Numero = " & NumeroFlete & " " _
       & "AND C.Codigo = TF.CodigoC " _
       & "AND CP.Item = TF.Item " _
       & "AND CP.Periodo = TF.Periodo " _
       & "AND CP.Codigo_Inv = TF.Codigo_Inv "
  Select_Adodc_Grid DGFletes, AdoFletes, sSQL
  With AdoFletes.Recordset
   If .RecordCount > 0 Then
       MBFechaI = .fields("Fecha_I")
       MBFechaF = .fields("Fecha_F")
       CodigoInv = .fields("Codigo_Inv")
       DCFacturaNo = .fields("Numero")
       TxtCarga = .fields("Carga")
       TxtCond = .fields("Referencia")
       TxtFlete = .fields("Flete")
       TxtPlaca = .fields("Ruta")
       TxtKmI = .fields("Km_Inicial")
       TxtKmF = .fields("Km_Final")
       CodigoCliente = .fields("CodigoC")
       DCCliente = .fields("Cliente")
       DCDestino = .fields("Producto")
       
       Codigo1 = .fields("Ayudante")
       Codigo2 = .fields("Conductor")
       DCAyudante = Ninguno
       DCConductor = Ninguno
       If AdoClientes.Recordset.RecordCount Then
          AdoClientes.Recordset.MoveFirst
          AdoClientes.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
          If Not AdoClientes.Recordset.EOF Then DCAyudante = AdoClientes.Recordset.fields("Cliente")
        
          AdoClientes.Recordset.MoveFirst
          AdoClientes.Recordset.Find ("Codigo = '" & Codigo2 & "' ")
          If Not AdoClientes.Recordset.EOF Then DCConductor = AdoClientes.Recordset.fields("Cliente")
       End If
       Label7.Caption = Format$(.fields("Factura"), "0000000")
       If Val(Label7.Caption) > 0 Then CommandButton1.Enabled = False Else CommandButton1.Enabled = True
   End If
  End With
  
  sSQL = "UPDATE Trans_Fletes " _
       & "SET Nombre_Ayudante = '.', Nombre_Conductor = '.' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
'  Ejecutar_SQL_SP sSQL
          
  If SQL_Server Then
     sSQL = "UPDATE Trans_Fletes " _
          & "SET Nombre_Ayudante = C.Cliente " _
          & "FROM Trans_Fletes As TF,Clientes As C "
  Else
     sSQL = "UPDATE Trans_Fletes As TF,Clientes As C " _
          & "SET TF.Nombre_Ayudante = C.Cliente "
  End If
  sSQL = sSQL & "WHERE TF.Item = '" & NumEmpresa & "' " _
       & "AND TF.Periodo = '" & Periodo_Contable & "' "
  sSQL = sSQL & "AND TF.Ayudante = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Trans_Fletes " _
          & "SET Nombre_Conductor = C.Cliente " _
          & "FROM Trans_Fletes As TF,Clientes As C "
  Else
     sSQL = "UPDATE Trans_Fletes As TF,Clientes As C " _
          & "SET TF.Nombre_Conductor = C.Cliente "
  End If
  sSQL = sSQL & "WHERE TF.Item = '" & NumEmpresa & "' " _
       & "AND TF.Periodo = '" & Periodo_Contable & "' "
''' If CodigoC <> "" Then
'''     sSQL = sSQL & "AND TF.CodigoC = '" & CodigoC & "' "
'''  Else
'''     sSQL = sSQL & "AND TF.CodigoC = '" & CodigoCliente & "' "
'''  End If
  sSQL = sSQL & "AND TF.Conductor = C.Codigo "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT Ok,Cliente,Numero,Fecha_I,Fecha_F,Factura,Producto As Placa,Carga,Flete,Km_Inicial,Km_Final,(Km_Final-Km_Inicial) As Recorrido,TF.Referencia,TF.Ruta,TF.CodigoC,TF.Nombre_Ayudante,TF.Nombre_Conductor " _
       & "FROM Clientes As C,Catalogo_Productos As CP,Trans_Fletes As TF " _
       & "WHERE TF.Item = '" & NumEmpresa & "' " _
       & "AND TF.Periodo = '" & Periodo_Contable & "' "
  If CodigoC <> "" Then
     sSQL = sSQL & "AND TF.CodigoC = '" & CodigoC & "' "
  Else
     sSQL = sSQL & "AND TF.CodigoC = '" & CodigoCliente & "' "
  End If
  sSQL = sSQL & "AND TF.T = '" & Normal & "' " _
       & "AND C.Codigo = TF.CodigoC " _
       & "AND CP.Item = TF.Item " _
       & "AND CP.Periodo = TF.Periodo " _
       & "AND CP.Codigo_Inv = TF.Codigo_Inv " _
       & "ORDER BY Cliente,Numero,Fecha_I "
  Select_Adodc_Grid DGFletes, AdoFletes, sSQL
  MBFechaI.SetFocus
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub TxtKmF_LostFocus()
   Label16.Caption = Format$(Val(TxtKmF) - Val(TxtKmI), "#,##0.00")
End Sub
