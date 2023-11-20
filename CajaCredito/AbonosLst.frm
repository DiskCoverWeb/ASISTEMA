VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form ListarAbonosPrest 
   Caption         =   "CANCELACION DE CREDITOS"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   10080
      TabIndex        =   40
      Top             =   105
      Visible         =   0   'False
      Width           =   1380
      Begin VB.OptionButton OpcC 
         Caption         =   "Caja"
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
         Left            =   210
         TabIndex        =   42
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OpcL 
         Caption         =   "Libreta"
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
         Left            =   210
         TabIndex        =   41
         Top             =   525
         Width           =   960
      End
   End
   Begin MSDataListLib.DataList DLTipoPrestamo 
      Bindings        =   "AbonosLst.frx":0000
      DataSource      =   "AdoTipoPrest"
      Height          =   2010
      Left            =   105
      TabIndex        =   3
      Top             =   420
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   3545
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo DCCreditos 
      Bindings        =   "AbonosLst.frx":001B
      DataSource      =   "AdoCreditos"
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   105
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "009000000"
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
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   330
      Left            =   1260
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA - <Crtl-B>: Buscar Pago por Número de Libretas"
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   192
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "AbonosLst.frx":0035
      Height          =   2640
      Left            =   105
      TabIndex        =   39
      Top             =   4200
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   4657
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   315
      Top             =   4305
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
      Caption         =   "CtaNo"
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
   Begin VB.TextBox TextLinea 
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
      Left            =   9030
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   1995
      Width           =   960
   End
   Begin VB.TextBox TxtNombresS 
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
      MaxLength       =   30
      TabIndex        =   8
      Top             =   735
      Width           =   5475
   End
   Begin VB.TextBox TextDias 
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
      Left            =   7560
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1995
      Width           =   1485
   End
   Begin VB.TextBox TextMeses 
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
      Left            =   6720
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1995
      Width           =   855
   End
   Begin VB.TextBox TextTasa 
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
      Left            =   5985
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   1995
      Width           =   750
   End
   Begin VB.TextBox TextSaldo 
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
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   2835
      Width           =   1695
   End
   Begin VB.TextBox TextCapital 
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
      Left            =   6510
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   2835
      Width           =   1590
   End
   Begin VB.TextBox TextComision 
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
      Left            =   5040
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   2835
      Width           =   1485
   End
   Begin VB.TextBox TextInt 
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
      Left            =   3465
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2835
      Width           =   1590
   End
   Begin VB.TextBox TextMonto 
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
      Left            =   1890
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2835
      Width           =   1590
   End
   Begin VB.TextBox TextNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
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
      Left            =   4515
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "AbonosLst.frx":004E
      Top             =   1995
      Width           =   1485
   End
   Begin VB.TextBox TextSaldoDisp 
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
      Left            =   105
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   2835
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
      Height          =   645
      Left            =   105
      MaxLength       =   100
      TabIndex        =   27
      Top             =   3465
      Width           =   9675
   End
   Begin VB.TextBox TextTP 
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
      Left            =   5775
      MaxLength       =   30
      TabIndex        =   15
      Top             =   1365
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
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
      Height          =   960
      Left            =   10080
      Picture         =   "AbonosLst.frx":0052
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1050
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   4515
      TabIndex        =   6
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1365
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   2415
      Top             =   4935
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
      Caption         =   "TipoPrest"
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
      Left            =   2415
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   315
      Top             =   5250
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
      Caption         =   "Asiento"
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
   Begin MSAdodcLib.Adodc AdoGarantes 
      Height          =   330
      Left            =   315
      Top             =   4935
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
      Caption         =   "Garantes"
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
   Begin MSAdodcLib.Adodc AdoTabla 
      Height          =   330
      Left            =   315
      Top             =   4620
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
      Caption         =   "Tabla"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   2415
      Top             =   5250
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
      Caption         =   "Creditos"
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
   Begin MSAdodcLib.Adodc AdoPrestamos 
      Height          =   330
      Left            =   2415
      Top             =   4305
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
      Caption         =   "Prestamos"
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Cuenta No."
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
      Width           =   1170
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " REPRESENTANTE DE LA CUENTA"
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
      TabIndex        =   7
      Top             =   420
      Width           =   5475
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &LIQUIDACION DE PRESTAMOS"
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
      TabIndex        =   4
      Top             =   105
      Width           =   5475
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Línea No"
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
      TabIndex        =   37
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Días Excedidos"
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
      Left            =   7560
      TabIndex        =   23
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuota #"
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
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tasa"
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
      Left            =   5985
      TabIndex        =   25
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo"
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
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Capital"
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
      Left            =   6510
      TabIndex        =   19
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision"
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
      TabIndex        =   26
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interes"
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
      Left            =   3465
      TabIndex        =   17
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto a pagar"
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
      Left            =   1890
      TabIndex        =   11
      Top             =   2520
      Width           =   1590
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Numero"
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
      TabIndex        =   35
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label LabelEgresos 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8295
      TabIndex        =   31
      Top             =   6930
      Width           =   1905
   End
   Begin VB.Label LabelIngresos 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6405
      TabIndex        =   32
      Top             =   6930
      Width           =   1905
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTALES"
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
      TabIndex        =   33
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Disponible"
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
      TabIndex        =   30
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Concepto"
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
      TabIndex        =   28
      Top             =   3150
      Width           =   9675
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE PRESTAMO"
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
      Left            =   5775
      TabIndex        =   9
      Top             =   1050
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA"
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
      TabIndex        =   5
      Top             =   1050
      Width           =   1275
   End
End
Attribute VB_Name = "ListarAbonosPrest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InsertarMontosPrestamo(DtaCta As Adodc, _
                                  CuentaNo As String, _
                                  TDebe As Currency, _
                                  THaber As Currency)
  If CuentaNo <> "00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
  If Si_No Then
     If OpcC.value Then
        sSQL = "SELECT TOP 1 * FROM Trans_Cajas " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' "
     Else
        sSQL = "SELECT TOP 1 * FROM Trans_Libretas " _
             & "WHERE Cuenta_No = '" & CuentaNo & "' " _
             & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
     End If
  Else
     sSQL = "SELECT TOP 1 * FROM Trans_Libretas " _
          & "WHERE Cuenta_No = '" & CuentaNo & "' " _
          & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  End If
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       SaldoDisp = 0: SaldoCont = 0
       ID_Trans = 0
       If Si_No Then
          If OpcL.value Then
             If .RecordCount > 0 Then
                 SaldoDisp = .Fields("Saldo_Disp")
                 SaldoCont = .Fields("Saldo_Cont")
                 ID_Trans = .Fields("IDT")
             End If
          End If
       Else
          If .RecordCount > 0 Then
              SaldoDisp = .Fields("Saldo_Disp")
              SaldoCont = .Fields("Saldo_Cont")
              ID_Trans = .Fields("IDT")
          End If
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Cuenta_No") = CuentaNo
       If Si_No Then
          If OpcC.value Then
            .Fields("TP") = "BOVE"
          Else
            .Fields("TP") = TipoProc
            .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
            .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
            .Fields("IDT") = ID_Trans + 1
          End If
       Else
         .Fields("TP") = TipoProc
         .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
         .Fields("IDT") = ID_Trans + 1
       End If
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = False
      .Fields("Cheque") = Ninguno
       SetUpdate DtaCta
  End With
  End If
End Sub

Public Sub ListarCuenta(Cuenta_No As String)
   TxtNombresS.Text = ""
   TxtNombresS.Text = ""
   TextMonto.Text = "0.00"
   TextInt.Text = "0.00"
   TextComision.Text = "0.00"
   TextCapital.Text = "0.00"
   TextSaldo.Text = "0.00"
   TextMeses.Text = "0"
   De_Vencidos = False: TotalEncaje = 0
   SaldoDisp = 0: SaldoCont = 0
   sSQL = "SELECT * " _
        & "FROM Trans_Bloqueos " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "AND T = 'N' "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           TotalEncaje = TotalEncaje + .Fields("Valor")
          .MoveNext
        Loop
    End If
   End With
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Libretas " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
   SelectAdodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      SaldoDisp = AdoAux.Recordset.Fields("Saldo_Disp")
      SaldoCont = AdoAux.Recordset.Fields("Saldo_Cont")
      TextLinea.Text = AdoAux.Recordset.Fields("ID")
   End If
   De_Vencidos = False
   SaldoDisp = SaldoDisp - TotalEncaje
   TextSaldoDisp.Text = Format(SaldoDisp, "#,##0.00")
   sSQL = "SELECT * " _
        & "FROM Clientes_Datos_Extras " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
        & "AND Tipo_Dato = 'LIBRETAS' "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Moneda_US = False '.Fields("ME")
        CodigoCli = .Fields("Codigo")
    End If
   End With
   sSQL = "SELECT * " _
        & "FROM Clientes " _
        & "WHERE Codigo = '" & CodigoCli & "' "
   SelectAdodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        TxtNombresS.Text = .Fields("Cliente")
        sSQL = "SELECT * FROM Clientes_Datos_Extras " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Tipo_Dato = 'LIBRETAS' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        'SelectAdodc AdoGarantes, sSQL
        TextMonto.Text = ""
        TextInt.Text = ""
        TextCapital.Text = ""
        TextSaldo.Text = ""
        TextMeses.Text = ""
        'TextNumero.Text = Format(Numero, "000000")
        sSQL = "SELECT * " _
             & "FROM Trans_Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND TP = '" & Codigo & "' " _
             & "AND Fecha = #" & BuscarFecha(Mifecha) & "# " _
             & "AND T = 'P' " _
             & "ORDER BY T,TP,Credito_No,Fecha "
        SelectAdodc AdoTabla, sSQL
        With AdoTabla.Recordset
         If .RecordCount > 0 Then
             De_Vencidos = .Fields("V")
             MBoxFecha.Text = .Fields("Fecha")
             TextMonto.Text = Format(.Fields("Pagos"), "#,##0.00")
             TextInt.Text = Format(.Fields("Interes"), "#,##0.00")
             TextComision.Text = Format(.Fields("Comision"), "#,##0.00")
             TextCapital.Text = Format(.Fields("Capital"), "#,##0.00")
             TextSaldo.Text = Format(.Fields("Saldo"), "#,##0.00")
             Total_Saldos = .Fields("Saldo")
             TotalCapital = .Fields("Capital")
             TotalInteres = .Fields("Interes") + .Fields("Comision")
             TextMeses.Text = .Fields("Cuota_No")
         End If
        End With
        Tasa = 0
        sSQL = "SELECT * " _
             & "FROM Prestamos " _
             & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Credito_No = '" & Contrato_No & "' " _
             & "AND TP = '" & Codigo & "' "
        SelectAdodc AdoTabla, sSQL
        If AdoTabla.Recordset.RecordCount > 0 Then Tasa = AdoTabla.Recordset.Fields("Tasa")
        TextTasa.Text = Format(Tasa, "00.00")
    End If
   End With
   Codigo = SinEspaciosIzq(DLTipoPrestamo.Text)
   With AdoPrestamos.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("CTP = '" & Codigo & "' ")
        If Not .EOF Then
           Si_No = .Fields("DM")
           TipoProc = .Fields("CTP")
           TextTP.Text = Codigo & "  " & .Fields("Descripcion")
           If Si_No Then Label5.Caption = " Dias" Else Label5.Caption = " Meses"
           Una_Vez = .Fields("DM")
        End If
    End If
   End With
End Sub

Private Sub Command3_Click()
 Unload ListarAbonosPrest
End Sub

Private Sub DCCreditos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCreditos_LostFocus()
  ListarCreditoPend
  DLTipoPrestamo.SetFocus
End Sub

Private Sub DLTipoPrestamo_DblClick()
  SiguienteControl
End Sub

Private Sub DLTipoPrestamo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLTipoPrestamo_LostFocus()
  Trans_No = 50
  IniciarAsientosDe DGAsiento, AdoAsiento
  If AdoTipoPrest.Recordset.RecordCount > 0 Then
     'AdoTipoPrest.Recordset.MoveFirst
     'DLTipoPrestamo.Text = AdoTipoPrest.Recordset.Fields("TipoP")
     Codigo = SinEspaciosIzq(DLTipoPrestamo.Text)
     Contrato_No = SinEspaciosDer(DLTipoPrestamo.Text)
     Mifecha = SinEspaciosIzqNoBlancos(DLTipoPrestamo.Text, 3)
     Cuenta_No = SinEspaciosIzqNoBlancos(DLTipoPrestamo.Text, 2)
     If Cuenta_No = "0" Then Cuenta_No = "00000000-0"
     If Codigo = "" Then Codigo = Ninguno
     If Mifecha = "datos." Then Mifecha = FechaSistema
 'Listar La Cuenta
  '
  'MsgBox MiFecha
  ListarCuenta Cuenta_No
  TextNumero.Text = Contrato_No
  Titulo = "TIPO DE TRANSACCION"
  Mensajes = "Transaccion en: [Si] Caja y [No] Libreta"
  If BoxMensaje = vbYes Then
     OpcC.value = True
     Si_No = True
  Else
     OpcL.value = True
     Si_No = False
  End If
  Frame1.Visible = True
  
  TextConcepto.Text = "(" & NumEmpresa & ") Abono No. " & TextMeses.Text & ", de Credito No. " & Contrato_No & ", Cuenta No. " & Cuenta_No & " de " & TxtNombresS.Text
  Total = Round(CCur(TextMonto.Text), 2)
  NoMeses = Round(CInt(TextMeses.Text), 2)
  Total_Interes = Round(CCur(TextInt.Text), 2)
  Total_Comision = Round(CCur(TextComision.Text), 2)
  NoDias = CFechaLong(FechaSistema) - CFechaLong(Mifecha)
  TextDias.Text = NoDias
  Haber = 0: Debe = CCur(TextCapital): Comision = 0
 'Asiento de Pago
  With AdoPrestamos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CTP = '" & TipoProc & "' ")
       If Not .EOF Then
          Una_Vez = .Fields("DM")
          If Una_Vez Then   ' Prestamos en dias
             If NoDias >= 1 Then
                Total = Round(CCur(TextMonto.Text), 2)
                Haber = Round(((Total * 0.05) / 30) * NoDias, 2)
             End If
          Else              'Prestamos en meses
            'Mas de 15 dias de mora
             sSQL = "SELECT * " _
                  & "FROM Tabla_Cobranzas " _
                  & "WHERE Cuota_hasta <> 0 " _
                  & "ORDER BY Cuota_hasta "
             SelectAdodc AdoCtaNo, sSQL
             If AdoCtaNo.Recordset.RecordCount > 0 Then
               'En caso que sea menos de quince dias de mora
                Comision = AdoCtaNo.Recordset.Fields("Dias_30") / 100
             End If
             If NoDias > 15 Then
                If AdoCtaNo.Recordset.RecordCount > 0 Then
                   SaldoInic = 0
                   Do While Not AdoCtaNo.Recordset.EOF
                      SaldoFinal = AdoCtaNo.Recordset.Fields("Cuota_hasta")
                      If SaldoInic <= Total And Total <= SaldoFinal Then
                         If 16 <= NoDias And NoDias <= 30 Then
                            Comision = AdoCtaNo.Recordset.Fields("Dias_30") / 100
                         ElseIf 31 <= NoDias And NoDias <= 60 Then
                            Comision = AdoCtaNo.Recordset.Fields("Dias_60") / 100
                         ElseIf 61 <= NoDias And NoDias <= 90 Then
                            Comision = AdoCtaNo.Recordset.Fields("Dias_90") / 100
                         Else
                            Comision = AdoCtaNo.Recordset.Fields("Porc_C") / 100
                         End If
                      End If
                      SaldoInic = SaldoFinal
                      AdoCtaNo.Recordset.MoveNext
                   Loop
                End If
             End If
          End If
          Haber = Round((Total * NoDias * Comision) / 360, 2)
        ' Asiento para la Libreta
          Total_Interes_Mora = Haber
          InsertarAsientos AdoAsiento, .Fields("Cta_Int_Mora"), 0, 0, Haber
          InsertarAsientos AdoAsiento, .Fields("Cta_Prestamo"), 0, 0, Debe
          InsertarAsientos AdoAsiento, .Fields("Cta_Interes"), 0, 0, Total_Interes
          InsertarAsientos AdoAsiento, .Fields("Cta_Comision"), 0, 0, Total_Comision
        ' Asiento de Efectivizacion de Prestamos
          InsertarAsientos AdoAsiento, .Fields("Cta_Int_Efec"), 0, Total_Interes, 0
          InsertarAsientos AdoAsiento, .Fields("Cta_Com_Efec"), 0, Total_Comision, 0
          InsertarAsientos AdoAsiento, .Fields("Cta_Int_Ganado"), 0, 0, Total_Interes
          InsertarAsientos AdoAsiento, .Fields("Cta_Com_Ganado"), 0, 0, Total_Comision
       End If
   End If
  End With
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  TotalLibreta = Haber - Debe
  Contra_Cta = Ninguno
  If Si_No Then
     InsertarAsientos AdoAsiento, Cta_CajaG, 0, TotalLibreta, 0
     Contra_Cta = Cta_CajaG
  Else
     InsertarAsientos AdoAsiento, Cta_Libretas, 0, TotalLibreta, 0
     Contra_Cta = Cta_Libretas
  End If
  Debe = 0: Haber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Debe = Debe + .Fields("DEBE")
          Haber = Haber + .Fields("HABER")
         .MoveNext
       Loop
   End If
  End With
  LabelIngresos.Caption = Format(Debe, "#,##0.00")
  LabelEgresos.Caption = Format(Haber, "#,##0.00")
  TxtNombresS.SetFocus
  Else
    MBoxCuenta.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
        & "FROM Catalogo_Prestamo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND TC <> " & Val(adFalse) & " " _
        & "ORDER BY CTP DESC "
  SelectAdodc AdoPrestamos, sSQL

  Mifecha = BuscarFecha(FechaSistema)
  TipoDoc = CompDiario
  Trans_No = 50
  IniciarAsientosDe DGAsiento, AdoAsiento
  ListarCreditoPend
  MBoxCuenta.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
   'CentrarForm Aprobacion
   ConectarAdodc AdoAux
   ConectarAdodc AdoCtaNo
   ConectarAdodc AdoTabla
   ConectarAdodc AdoAsiento
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoGarantes
   ConectarAdodc AdoPrestamos
   ConectarAdodc AdoTipoPrest
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub ListarCreditoPend()
  'Si_No = OpcL.Value
  Credito_No = SinEspaciosDer(DCCreditos.Text)
  If SQL_Server Then
     sSQL = "SELECT (P.TP & ' ' & P.Cuenta_No & ' ' & CONVERT(nvarchar(10),P.Fecha,103 ) & ' ' & P.Credito_No) As TipoP "
  Else
     sSQL = "SELECT (P.TP & ' ' & P.Cuenta_No & ' ' & CSTR(P.Fecha) & ' ' & P.Credito_No) As TipoP "
  End If
  sSQL = sSQL & "FROM Trans_Prestamos As P " _
       & "WHERE P.Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND P.Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND P.Credito_No = '" & Credito_No & "' " _
       & "AND P.T = 'P' " _
       & "ORDER BY P.TP,P.Cuenta_No,P.Fecha,P.Credito_No "
  SelectDBList DLTipoPrestamo, AdoTipoPrest, sSQL, "TipoP"
End Sub

Private Sub MBoxCuenta_LostFocus()
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Moneda_US = False  '.Fields("ME")
       CodigoCli = .Fields("Codigo")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo = '" & CodigoCli & "' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then TxtNombresS.Text = .Fields("Cliente")
  End With
  sSQL = "SELECT TP & ' ' & Credito_No As TipoCred " _
       & "FROM Trans_Prestamos " _
       & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' " _
       & "AND T = 'P' " _
       & "GROUP BY TP,Credito_No "
  'MsgBox sSQL
  SelectDBCombo DCCreditos, AdoCreditos, sSQL, "TipoCred"
End Sub

