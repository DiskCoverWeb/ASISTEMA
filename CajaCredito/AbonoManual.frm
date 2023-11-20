VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form AbonoPrestamoManual 
   Caption         =   "CANCELACION DE CREDITOS"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
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
      Left            =   6195
      MaxLength       =   30
      TabIndex        =   15
      Top             =   2205
      Width           =   5055
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
      Left            =   4620
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2205
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
      Left            =   3150
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2205
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
      Left            =   1575
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2205
      Width           =   1590
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Abono"
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
      Left            =   8610
      Picture         =   "AbonoManual.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   105
      Width           =   1275
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
      Left            =   2415
      MaxLength       =   30
      TabIndex        =   5
      Top             =   420
      Width           =   6105
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
      Left            =   105
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "AbonoManual.frx":08CA
      Top             =   2205
      Width           =   1485
   End
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "AbonoManual.frx":08CE
      Height          =   2640
      Left            =   105
      TabIndex        =   23
      Top             =   2625
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
      Top             =   2940
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
      MaxLength       =   119
      TabIndex        =   17
      Top             =   1155
      Width           =   9780
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Pago"
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
      Left            =   9975
      Picture         =   "AbonoManual.frx":08E7
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   105
      Width           =   1275
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
      Height          =   645
      Left            =   9975
      Picture         =   "AbonoManual.frx":0D29
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      ToolTipText     =   "Formato de Fecha: DD/MM/AA - <Crtl-B>: Buscar Pago por Número de Libretas"
      Top             =   420
      Width           =   1170
      _ExtentX        =   2064
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   2415
      Top             =   3570
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
      Top             =   3255
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
      Top             =   3885
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
      Top             =   3570
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
      Top             =   3255
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
      Top             =   3885
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
      Top             =   2940
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
   Begin MSMask.MaskEdBox MBFechaP 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
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
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
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
      Left            =   6195
      TabIndex        =   14
      Top             =   1890
      Width           =   5055
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
      Left            =   4620
      TabIndex        =   12
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Mora"
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
      TabIndex        =   10
      Top             =   1890
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
      Left            =   1575
      TabIndex        =   8
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Pago"
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
      Left            =   2415
      TabIndex        =   4
      Top             =   105
      Width           =   6105
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuota No."
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
      Top             =   1890
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
      Left            =   9030
      TabIndex        =   20
      Top             =   5355
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
      Left            =   7140
      TabIndex        =   21
      Top             =   5355
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
      Left            =   6090
      TabIndex        =   22
      Top             =   5355
      Width           =   1065
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
      TabIndex        =   16
      Top             =   840
      Width           =   9780
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
      Left            =   1260
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "AbonoPrestamoManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
FechaValida MBFechaP
If (Total_Interes + TotalCapital + Total_Comision) > 0 Then
Titulo = "Pregunta de Grabacion"
Mensajes = "Seguro de Grabar Abono"
If BoxMensaje = vbYes Then
   RatonReloj
   FechaTexto = MBFechaP
   If Round(SumaDebe - SumaHaber, 2) = 0 Then
      Co.T = Normal
      Co.TP = CompIngreso
      Co.Fecha = FechaTexto
      Co.CodigoB = CodigoCli
      Co.Efectivo = TotalLibreta
      Co.Monto_Total = TotalLibreta
      Co.Numero = ReadSetDataNum("Ingresos", True, True)
      Co.Concepto = TextConcepto.Text
      Co.Item = NumEmpresa
      Co.T_No = Trans_No
      Co.Usuario = CodigoUsuario
      GrabarComprobante Co
      
      SetAdoAddNew "Trans_Libretas"
      SetAdoFields "T", Normal
      SetAdoFields "ME", False
      SetAdoFields "Fecha", FechaTexto
      SetAdoFields "Cuenta_No", "BOVEDA"
      SetAdoFields "TP", "BOVE"
      SetAdoFields "Debitos", 0
      SetAdoFields "Creditos", TotalLibreta
      SetAdoFields "CodigoU", CodigoUsuario
      SetAdoFields "Hora", Format(Time, FormatoTimes)
      SetAdoFields "Item", NumEmpresa
      SetAdoFields "Cheque", Ninguno
      SetAdoFields "ACC", False
      SetAdoFields "CHT", False
      SetAdoUpdate
      ImprimirComprobantesDe False, Co
      ImprimirComprobantesDe False, Co
   End If
   Mifecha = BuscarFecha(FechaTexto)
   TipoDoc = CompDiario
   Trans_No = 50
   IniciarAsientosDe DGAsiento, AdoAsiento
   MBoxCuenta.SetFocus
   RatonNormal
End If
Else
   MsgBox "No se puede grabar el abono, faltan datos"
End If
End Sub

Private Sub Command3_Click()
 Unload AbonoPrestamoManual
End Sub

Private Sub Command4_Click()
On Error GoTo Errorhandler
Titulo = "IMPRESIONES"
Mensajes = "Imprimir Comprobante de Pago"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   Cuenta_No = MBoxCuenta.Text
   TotalCapital = Val(CCur(TextCapital))
   TotalInteres = Val(CCur(TextInt))
   Total_Interes_Mora = Val(CCur(TextComision))
   TotalLibreta = TotalCapital + TotalInteres + Total_Interes_Mora
   
   InicioX = 0.5: InicioY = 0
   DataAnchoCampos InicioX, AdoAsiento, 8, TipoTimes, Orientacion_Pagina
   Pagina = 1
   EncabezadoEmpresa 0.1
   PrinterPaint LogoTipo, 2, 0.1, 3, 1.5
   Printer.FontName = TipoTimes
   Printer.FontSize = 12: Printer.FontBold = True
   PrinterTexto 2, 2, "C O M P R O B A N T E   D E   P A G O"
   Printer.FontSize = 11
   PrinterTexto 12, 2, NombreCiudad & ", " & FechaStrg(MBFechaP)
   PrinterTexto 2, 2.6, "Abono del Préstamo"
   PrinterTexto 12, 2.6, "Cta. Ahorro No. " & Cuenta_No
   PrinterTexto 2, 3.2, UCase(TextTP.Text)
   PrinterTexto 12, 3.2, "Cuota No. " & TextNumero.Text
   PrinterTexto 2, 3.9, "SOCIO:"
   PrinterTexto 12, 3.9, "Capital"
   PrinterTexto 2, 4.5, "La cantidad de:"
   PrinterTexto 12, 4.4, "Interés"
   PrinterTexto 12, 4.9, "Interés Mora"
   PrinterTexto 12, 5.5, "Total Abono"
   PrinterTexto 14.8, 3.9, Moneda
   PrinterTexto 14.8, 4.4, Moneda
   PrinterTexto 14.8, 4.9, Moneda
   PrinterTexto 14.8, 5.5, Moneda
   Printer.FontBold = False
   PrinterTexto 3.5, 3.9, TxtNombresS.Text
   PrinterLineas 4.8, 4.5, Cambio_Letras(TotalLibreta), 7, 0.45
   PrinterVariables 16, 3.9, TotalCapital
   PrinterVariables 16, 4.4, TotalInteres
   PrinterVariables 16, 4.9, Total_Interes_Mora
   PrinterVariables 16, 5.5, TotalLibreta
   Imprimir_Linea_H 5.4, 12, 19, Negro, True
   PrinterTexto 2, 6, String(18, "_")
   PrinterTexto 8, 6, String(11, "_")
   PrinterTexto 2.1, 6.5, "Cajero: " & CodigoUsuario
   PrinterTexto 8.3, 6.5, "Conforme"
   RatonNormal
   Printer.EndDoc
   MensajeEncabData = ""
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
        & "FROM Catalogo_Prestamo " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND TC <> " & Val(adFalse) & " " _
        & "AND CTP = 'S/F'"
  SelectAdodc AdoPrestamos, sSQL

  Mifecha = BuscarFecha(FechaSistema)
  TipoDoc = CompDiario
  Trans_No = 50
  IniciarAsientosDe DGAsiento, AdoAsiento
  MBFechaP.SetFocus
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

Private Sub MBFechaP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaP_LostFocus()
  FechaValida MBFechaP
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_LostFocus()
  TxtNombresS = Ninguno
  CodigoCli = Ninguno
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
   If .RecordCount > 0 Then
       TxtNombresS.Text = .Fields("Cliente")
       TextTP.Text = "PRESTAMO ANTIGUO"
       TextNumero.SetFocus
   End If
  End With
End Sub

Private Sub TextCapital_GotFocus()
  MarcarTexto TextCapital
End Sub

Private Sub TextCapital_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCapital_LostFocus()
  Debe = 0: Haber = 0
  TextoValido TextCapital, True
  Cuenta_No = MBoxCuenta
  Trans_No = 50
  IniciarAsientosDe DGAsiento, AdoAsiento
  NoMeses = Round(CInt(TextNumero), 2)
  Total_Interes = Round(CCur(TextInt), 2)
  Total_Comision = Round(CCur(TextComision), 2)
  TotalCapital = Round(CCur(TextCapital), 2)
  TotalLibreta = Total_Interes + Total_Comision + TotalCapital
 'Asiento de Abonos prestamos antiguos
  TextConcepto = "(" & NumEmpresa & ") Abono No. " & TextNumero & ", Credito de la Cuenta No. " & Cuenta_No & " de " & TxtNombresS & ", Por USD " & Format(TotalLibreta, "#,##0.00")
  If (Total_Interes >= 0) Or (TotalCapital >= 0) Or (Total_Comision >= 0) Then
     InsertarAsientos AdoAsiento, Cta_CajaG, 0, TotalLibreta, 0
     With AdoPrestamos.Recordset
      If .RecordCount > 0 Then
          'MsgBox ".. "
          InsertarAsientos AdoAsiento, .Fields("Cta_Prestamo"), 0, 0, TotalCapital
          InsertarAsientos AdoAsiento, .Fields("Cta_Int_Ganado"), 0, 0, Total_Interes
          InsertarAsientos AdoAsiento, .Fields("Cta_Int_Mora"), 0, 0, Total_Comision
      End If
     End With
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
     Command2.SetFocus
  Else
     MsgBox "Falta de ingresar mas datos"
     LabelIngresos.Caption = Format(Debe, "#,##0.00")
     LabelEgresos.Caption = Format(Haber, "#,##0.00")
     TextNumero.SetFocus
  End If
  
End Sub

Private Sub TextComision_GotFocus()
  MarcarTexto TextComision
End Sub

Private Sub TextComision_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextComision_LostFocus()
  TextoValido TextComision, True
End Sub

Private Sub TextInt_GotFocus()
  MarcarTexto TextInt
End Sub

Private Sub TextInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInt_LostFocus()
  TextoValido TextInt, True
End Sub

Private Sub TextNumero_GotFocus()
  MarcarTexto TextNumero
End Sub

Private Sub TextNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextNumero_LostFocus()
  TextoValido TextNumero, True
  TextNumero = Format(TextNumero, "#,##0")
End Sub
