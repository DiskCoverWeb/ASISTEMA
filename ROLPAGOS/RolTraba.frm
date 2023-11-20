VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FRolMinisterioTrabajo 
   BackColor       =   &H8000000D&
   Caption         =   "BANCO BOLIVARIANO"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14640
   Icon            =   "RolTraba.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "RolTraba.frx":0442
   ScaleHeight     =   7530
   ScaleWidth      =   14640
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OpcT 
      Caption         =   "Todos"
      Height          =   225
      Left            =   7140
      TabIndex        =   4
      Top             =   105
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.OptionButton OpcF 
      Caption         =   "Femenino"
      Height          =   225
      Left            =   7140
      TabIndex        =   6
      Top             =   735
      Width           =   1170
   End
   Begin VB.OptionButton OpcM 
      Caption         =   "Masculino"
      Height          =   225
      Left            =   7140
      TabIndex        =   5
      Top             =   420
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12180
      Picture         =   "RolTraba.frx":4F7A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar Decimo 3ro."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9660
      Picture         =   "RolTraba.frx":5284
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar Decimo 4to."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      Picture         =   "RolTraba.frx":5A8E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid DGClientes 
      Bindings        =   "RolTraba.frx":6298
      Height          =   1380
      Left            =   105
      TabIndex        =   12
      Top             =   1575
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   2434
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Height          =   855
      Left            =   13440
      Picture         =   "RolTraba.frx":62B2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar Contratos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      Picture         =   "RolTraba.frx":6CA8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4095
      Top             =   2520
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   1890
      Top             =   2520
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   4095
      Top             =   2835
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
      Caption         =   "Producto"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   4410
      TabIndex        =   1
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   1890
      Top             =   3150
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
      Caption         =   "Abono"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   5775
      TabIndex        =   3
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoPendiente 
      Height          =   330
      Left            =   1890
      Top             =   1890
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
      Caption         =   "Pendiente"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   6405
      Width           =   5160
      _ExtentX        =   9102
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   1890
      Top             =   2835
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
   Begin MSDataGridLib.DataGrid DGAsiento 
      Bindings        =   "RolTraba.frx":74B2
      Height          =   2640
      Left            =   105
      TabIndex        =   13
      Top             =   3675
      Width           =   10410
      _ExtentX        =   18362
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
   Begin MSAdodcLib.Adodc AdoTipoPrest 
      Height          =   330
      Left            =   3780
      Top             =   5040
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   3780
      Top             =   5355
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
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   330
      Left            =   4095
      Top             =   2205
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "RolTraba.frx":74CB
      DataSource      =   "AdoBanco"
      Height          =   345
      Left            =   2415
      TabIndex        =   18
      Top             =   1260
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Banco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   4095
      Top             =   1890
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
      Caption         =   "Banco"
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
   Begin VB.Label LblConcepto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
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
      Top             =   3360
      Width           =   10305
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta de Transferencia"
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
      Top             =   1260
      Width           =   2325
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
      TabIndex        =   16
      Top             =   6405
      Width           =   1065
   End
   Begin VB.Label LabelIngresos 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6405
      TabIndex        =   15
      Top             =   6405
      Width           =   1905
   End
   Begin VB.Label LabelEgresos 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   8295
      TabIndex        =   14
      Top             =   6405
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Final"
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
      TabIndex        =   2
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Inicio"
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
      Left            =   4410
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "FRolMinisterioTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim NumFileAlumnos As Integer
Dim RutaGeneraFileAlumnos As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal
Dim RutaBackupXX As String
Dim CtasPro() As CtasAsiento
Dim CantCtas As Long

Private Sub Command1_Click()
  CaptionTemp = FRolMinisterioTrabajo.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  Trans_No = 100
  BorrarAsientos True
  IniciarAsientosDe DGAsiento, AdoAsiento
  Ctas_Asientos_Decimos
  TipoDoc = ""
  AuxNumEmp = NumEmpresa
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  FechaTexto = BuscarFecha(MBFechaF)
  FechaTexto1 = Format(MBFechaF, "MM/dd/yyyy")
  Presentar_Rol_Pago 4
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
  Generar_Decimo 4
  Procesar_Asientos_Rol 4
  FRolMinisterioTrabajo.Caption = CaptionTemp
  LblConcepto.Caption = "Transferencia pago del Décimo Cuarto: " & MesesLetras(Month(MBFechaI)) & "-" & Year(MBFechaI)
End Sub

Private Sub Command2_Click()
  Unload FRolMinisterioTrabajo
End Sub

Private Sub Command3_Click()
  CaptionTemp = FRolMinisterioTrabajo.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  Trans_No = 100
  BorrarAsientos True
  IniciarAsientosDe DGAsiento, AdoAsiento
  Ctas_Asientos_Decimos
  TipoDoc = ""
  AuxNumEmp = NumEmpresa
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  FechaTexto = BuscarFecha(MBFechaF)
  FechaTexto1 = Format(MBFechaF, "MM/dd/yyyy")
  Presentar_Rol_Pago 3
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
  Generar_Decimo 3
  Procesar_Asientos_Rol 3
  FRolMinisterioTrabajo.Caption = CaptionTemp
  LblConcepto.Caption = "Transferencia pago del Décimo Tercero: " & MesesLetras(Month(MBFechaI)) & "-" & Year(MBFechaI)
End Sub

Private Sub Command4_Click()
Dim Cont As Integer
Dim CaptionTemp As String
  CaptionTemp = FRolMinisterioTrabajo.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  TipoDoc = ""
  AuxNumEmp = NumEmpresa
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  FechaTexto = BuscarFecha(MBFechaF)
  FechaTexto1 = Format(MBFechaF, "MM/dd/yyyy")
 'Detalle de las Facturas Emitidas del mes
  sSQL = "SELECT C.CI_RUC, C.Cliente, CR.Codigo " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' "
  If OpcF.value Then
     sSQL = sSQL & "AND C.Sexo = 'F' "
  ElseIf OpcM.value Then
     sSQL = sSQL & "AND C.Sexo = 'M' "
  End If
  sSQL = sSQL & "AND CR.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente,CR.Codigo "
  Select_Adodc_Grid DGClientes, AdoClientes, sSQL
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
  Generar_Ministerio
  FRolMinisterioTrabajo.Caption = CaptionTemp
  Co.Concepto = ""
End Sub

Private Sub Command5_Click()
  Trans_No = 100
  FechaComp = MBFechaI
  FechaValida MBFechaI: FechaIni = BuscarFecha(MBFechaI)
  FechaValida MBFechaF: FechaFin = BuscarFecha(MBFechaF)
  SumaDebe = 0: SumaHaber = 0
  With AdoAsiento.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaDebe = SumaDebe + .Fields("DEBE")
          SumaHaber = SumaHaber + .Fields("HABER")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
 'MsgBox SumaDebe & vbCrLf & SumaHaber
  If Redondear(SumaDebe - SumaHaber, 2) = 0 Then
     RatonReloj
     Co.T = Normal
     Co.Fecha = MBFechaF
     Co.CodigoB = Ninguno
     Co.Efectivo = SumaHaber
     Co.Monto_Total = SumaHaber
     Co.Item = NumEmpresa
     Co.Usuario = CodigoUsuario
     Co.TP = CompDiario
     Co.Numero = ReadSetDataNum("Diario", True, True)
     Co.Concepto = LblConcepto.Caption
     Co.T_No = Trans_No
     GrabarComprobante Co
     ImprimirComprobantesDe False, Co
  Else
     MsgBox "No se puede grabar, descuadre en el asiento"
  End If
End Sub

Private Sub DGClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then GenerarDataTexto FRolMinisterioTrabajo, AdoClientes, True
  If CtrlDown And (vbKeyP = KeyCode) Then ImprimirAdodc AdoClientes, 2, 7, True
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('CJ','BA') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC,Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"

  Trans_No = 100
  BorrarAsientos True
  IniciarAsientosDe DGAsiento, AdoAsiento
  DGClientes.Caption = "ORIGEN" & Space(18) & "COD: " & CodigoDelBanco
  RutaOrigen = RutaSistema & "\LOGOS\MINISTER.GIF"
  'FRolMinisterioTrabajo.BackColor = &H80FFFF
  FRolMinisterioTrabajo.Caption = "ARCHIVO MAESTRO DEL MINISTERIO DE TRABAJO (" & CodigoDelBanco & ")"
  FRolMinisterioTrabajo.Picture = LoadPicture(RutaOrigen)
  Label1.BackColor = FRolMinisterioTrabajo.BackColor
  
  Label8.BackColor = FRolMinisterioTrabajo.BackColor
  MBFechaI.BackColor = FRolMinisterioTrabajo.BackColor
  MBFechaF.BackColor = FRolMinisterioTrabajo.BackColor
  OpcM.BackColor = FRolMinisterioTrabajo.BackColor
  OpcF.BackColor = FRolMinisterioTrabajo.BackColor
  OpcT.BackColor = FRolMinisterioTrabajo.BackColor
  RatonNormal
End Sub

Private Sub Form_Deactivate()
  FRolMinisterioTrabajo.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
Dim Mitad As Single
Dim alto As Single
 'CentrarForm FRolMinisterioTrabajo
  alto = MDI_Y_Max / 2
  RutaBackupXX = ""
  ConectarAdodc AdoAux
  ConectarAdodc AdoAbono
  ConectarAdodc AdoBanco
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoFactura
  ConectarAdodc AdoClientes
  ConectarAdodc AdoProducto
  ConectarAdodc AdoPendiente
  ConectarAdodc AdoAsiento
  DGClientes.width = MDI_X_Max - DGClientes.Left - 100
  DGAsiento.width = MDI_X_Max - DGClientes.Left - 100
  LblConcepto.width = MDI_X_Max - DGClientes.Left - 100
  DCBanco.width = MDI_X_Max - DCBanco.Left - 100
  DGClientes.Height = alto
  LblConcepto.Top = DGClientes.Top + DGClientes.Height
  DGAsiento.Top = LblConcepto.Top + LblConcepto.Height
  DGAsiento.Height = MDI_Y_Max - DGAsiento.Top - 300
  Label17.Top = DGAsiento.Top + DGAsiento.Height
  LabelIngresos.Top = DGAsiento.Top + DGAsiento.Height
  LabelEgresos.Top = DGAsiento.Top + DGAsiento.Height
  AdoClientes.Top = DGAsiento.Top + DGAsiento.Height
  RutaBackup = RutaSysBases & "\BANCO"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  MBFechaF.Text = UltimoDiaMes(MBFechaI.Text)
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

Public Sub Generar_Ministerio()
Dim AuxNumEmp As String

Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Traza = Year(MBFechaI) & "-" & UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3)) & "-" & Format(Day(MBFechaI), "00")
  RutaGeneraFileAlumnos = UCase(RutaBackup & "\RolPago\" & UCase(Empresa) & "_" & Traza) & ".TXT"
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 0
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If Len(.Fields("CI_RUC")) = 10 Then
             Contador = Contador + 1
             NombreCliente = ULCase(.Fields("Cliente"))
             Codigo1 = SinEspaciosIzq(NombreCliente)
             NombreCliente = TrimStrg(MidStrg(NombreCliente, Len(Codigo1) + 1, Len(NombreCliente)))
             Codigo2 = SinEspaciosIzq(NombreCliente)
             NombreCliente = TrimStrg(MidStrg(NombreCliente, Len(Codigo2) + 1, Len(NombreCliente)))
             Codigo3 = NombreCliente
             If Len(Codigo3) <= 1 Then Codigo3 = ""
             CodigoP = .Fields("CI_RUC")
           ' Empieza la trama por Alumno
             Print #NumFileAlumnos, CodigoP & ";";
             Print #NumFileAlumnos, TrimStrg(Codigo3) & ";";
             Print #NumFileAlumnos, TrimStrg(Codigo1 & " " & Codigo2) & ";"
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
 'Finalizamos los Archivos
  RatonNormal
  MsgBox "SE GENERO EL SIGUIENTE ARCHIVO: " & vbCrLf & vbCrLf & vbCrLf _
       & RutaGeneraFileAlumnos & vbCrLf & vbCrLf
End Sub

Public Sub Generar_Decimo(Decimo As Byte)
Dim AuxNumEmp As String

Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
  RatonReloj
  Separador = ","
 'Separador = vbTab
 'Abrimo los archivo que vamos ha necesitar
  Select Case Decimo
    Case 3: Traza = "Decimo_3ro-" & Year(MBFechaI) & "-" & UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3))
    Case 4: Traza = "Decimo_4to-" & Year(MBFechaI) & "-" & UCase(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3))
  End Select
  RutaGeneraFileAlumnos = UCase(RutaBackup & "\RolPago\" & Traza) & ".csv"
  Traza = ""
  NumFileAlumnos = FreeFile
  Open RutaGeneraFileAlumnos For Output As #NumFileAlumnos
  FechaTexto = FechaSistema
  Mifecha = Format(Year(MBFechaP), "0000") _
          & Format(Month(MBFechaP), "00") _
          & Format(Day(MBFechaP), "00")
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 0
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Select Case Decimo
         Case 3: Print #NumFileAlumnos, "Cédula (Ejm.:0502366503);Nombres;Apellidos;Genero (Masculino=M ó Femenino=F);Ocupación;Total_ganado (Ejm.:1000.56);Días laborados (360 días equivalen a un año);Tipo de Deposito(Pago Directo=P,Acreditación en Cuenta=A,Retencion Pago Directo=RP,Retencion Acreditación en Cuenta=RA);Solo si el trabajador posee JORNADA PARCIAL PERMANENTE ponga una X, si tiene retencion ponga una R;DETERMINE EN HORAS LA JORNADA PARCIAL PERMANENTE SEMANAL ESTIPULADO EN EL CONTRATO;Solo si su trabajador posee algun tipo de discapacidad ponga una X;"
         Case 4: Print #NumFileAlumnos, "Cédula (Ejm.:0502366503);Nombres;Apellidos;Genero (Masculino=M ó Femenino=F);Ocupación(codigo iess);Días laborados (360 días equivalen a un año);Tipo de Pago(Pago Directo=P,Acreditación en Cuenta=A,Retencion Pago Directo=RP,Retencion Acreditación en Cuenta=RA);Solo si el trabajador posee JORNADA PARCIAL PERMANENTE ponga una X;DETERMINE EN HORAS LA JORNADA PARCIAL PERMANENTE SEMANAL ESTIPULADO EN EL CONTRATO;Solo si su trabajador posee algun tipo de discapacidad ponga una X;Fecha de Jubilación;valor Retencion;SOLO SI SU TRABAJADOR MENSUALIZA EL PAGO DE LA DECIMOCUARTA REMUNERACIÓN PONGA UNA X"
       End Select
       Do While Not .EOF
          If Len(.Fields("CI_RUC")) = 10 Then
             Contador = Contador + 1
             NombreCliente = UCase(.Fields("Cliente"))
             Codigo1 = SinEspaciosIzq(NombreCliente)
             NombreCliente = TrimStrg(MidStrg(NombreCliente, Len(Codigo1) + 1, Len(NombreCliente)))
             Codigo2 = SinEspaciosIzq(NombreCliente)
             NombreCliente = TrimStrg(MidStrg(NombreCliente, Len(Codigo2) + 1, Len(NombreCliente)))
             Codigo3 = NombreCliente
             If Len(Codigo3) <= 1 Then Codigo3 = ""
             CodigoP = .Fields("CI_RUC")
           ' Empieza la trama por Empleado
             Select Case Decimo
               Case 3
                     Print #NumFileAlumnos, CodigoP & ";";
                     Print #NumFileAlumnos, TrimStrg(Codigo3) & ";";
                     Print #NumFileAlumnos, TrimStrg(Codigo1 & " " & Codigo2) & ";";
                     Print #NumFileAlumnos, TrimStrg(.Fields("Sexo")) & ";";
                     Print #NumFileAlumnos, .Fields("Profesion") & ";";
                     Print #NumFileAlumnos, Format(.Fields("Valor_Dec_3ro"), "##0.00") & ";";
                     Print #NumFileAlumnos, .Fields("Dias_Dec_3ro") & ";";
                     Print #NumFileAlumnos, .Fields("FormaPago10to") & ";"
'''                     Print #NumFileAlumnos, ";";
'''                     Print #NumFileAlumnos, ";"
               Case 4
                     Print #NumFileAlumnos, CodigoP & ";";
                     Print #NumFileAlumnos, TrimStrg(Codigo3) & ";";
                     Print #NumFileAlumnos, TrimStrg(Codigo1 & " " & Codigo2) & ";";
                     Print #NumFileAlumnos, TrimStrg(.Fields("Sexo")) & ";";
                     Print #NumFileAlumnos, .Fields("Profesion") & ";";
                     Print #NumFileAlumnos, .Fields("Dias_Dec_4to") & ";";
                     Print #NumFileAlumnos, .Fields("FormaPago10to") & ";";
                     Print #NumFileAlumnos, ";";
                     Print #NumFileAlumnos, ";";
                     If .Fields("Porcentaje") > 0 Then
                         Print #NumFileAlumnos, "X;";
                     Else
                         Print #NumFileAlumnos, ";";
                     End If
                     Print #NumFileAlumnos, ";";
                     If .Fields("Pagar_Decimos") Then
                         Print #NumFileAlumnos, ";X"
                     Else
                         Print #NumFileAlumnos, ";"
                     End If
             End Select
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileAlumnos
' Comenzamos a generar el asiento contable
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoP = .Fields("CI_RUC")
             Select Case Decimo
               Case 3
                    InsValorCtaPro .Fields("Cta_Decimo_Tercer_P"), .Fields("Valor_Dec_3ro")
                   '.Fields ("Dias_Dec_3ro")
               Case 4
                    InsValorCtaPro .Fields("Cta_Decimo_Cuarto_P"), .Fields("Valor_Dec_4to")
                   '.Fields ("Dias_Dec_4to")
             End Select
         .MoveNext
       Loop
   End If
  End With
 'Finalizamos los Archivos
  RatonNormal
End Sub

Public Sub Presentar_Rol_Pago(TipoArchivo As Byte)
 'Detalle del Rol de Pago
 
  sSQL = "SELECT CR.Fecha As Fecha_Ing,C.CI_RUC,C.TD,C.Cliente,C.Sexo,C.Profesion,CR.Porcentaje," _
       & "CR.CodProfesion,CR.FormaPago10to,CR.FP,CR.Codigo," _
       & "CR.Valor_Dec_3ro,CR.Valor_Dec_4to," _
       & "CR.Dias_Dec_3ro,CR.Dias_Dec_4to," _
       & "Cta_Decimo_Tercer_P,Cta_Decimo_Cuarto_P,Pagar_Decimos " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' "
  Select Case TipoArchivo
   Case 3: sSQL = sSQL & "AND CR.Valor_Dec_3ro > 0 "
   Case 4: sSQL = sSQL & "AND CR.Valor_Dec_4to > 0 "
  End Select
  If OpcF.value Then
     sSQL = sSQL & "AND C.Sexo = 'F' "
  ElseIf OpcM.value Then
     sSQL = sSQL & "AND C.Sexo = 'M' "
  End If
  sSQL = sSQL & "AND CR.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente,CR.Codigo "
  Select_Adodc_Grid DGClientes, AdoClientes, sSQL
End Sub

Public Sub Ctas_Asientos_Decimos()
  RatonReloj
  sSQL = "SELECT Grupo_Rol, Cta_Decimo_Tercer_P, Cta_Decimo_Cuarto_P " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Grupo_Rol, Cta_Decimo_Tercer_P, Cta_Decimo_Cuarto_P "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       CantCtas = (.RecordCount * 2) + 10
       ReDim CtasPro(CantCtas) As CtasAsiento
       For IE = 0 To CantCtas - 1
           CtasPro(IE).Cta = "0"
           CtasPro(IE).Valor = 0
       Next IE
      'Seteamos las Cuentas del Rol Pagos
       Do While Not .EOF
         'Provisiones de Decimos
          SetearCtasCierrePro .Fields("Cta_Decimo_Tercer_P")
          SetearCtasCierrePro .Fields("Cta_Decimo_Cuarto_P")
         .MoveNext
       Loop
   End If
  End With
  Cta = SinEspaciosIzq(DCBanco)
  SetearCtasCierrePro Cta
  RatonNormal
End Sub

Public Sub SetearCtasCierrePro(CtaFields As String)
  Si_No = True
  For IE = 0 To CantCtas - 1
      If CtaFields = CtasPro(IE).Cta Then Si_No = False
  Next IE
  If Si_No Then
     IE = 0
     While IE < CantCtas
        If CtasPro(IE).Cta = "0" Then
           'MsgBox Leer_Cta_Catalogo(CtaFields)
           If Leer_Cta_Catalogo(CtaFields) <> Ninguno Then
              CtasPro(IE).Cta = CtaFields
              IE = CantCtas + 1
           End If
        End If
        IE = IE + 1
     Wend
  End If
End Sub

Public Sub Procesar_Asientos_Rol(Decimo As Byte)
Dim VentasDia As Boolean
Dim Ctas_Catalogo As String
Dim Total_Aporte_Patronal As Currency
   RatonReloj
   CodigoCli = Ninguno
   I = CantCtas - 1
   For IE = 0 To I - 1
     For JE = IE + 1 To I
       If CtasPro(IE).Cta < CtasPro(JE).Cta Then
          Cta_Aux = CtasPro(IE).Cta
          Valor = CtasPro(IE).Valor
          CtasPro(IE).Cta = CtasPro(JE).Cta
          CtasPro(IE).Valor = Redondear(CtasPro(JE).Valor, 2)
          CtasPro(JE).Cta = Cta_Aux
          CtasPro(JE).Valor = Valor
       End If
     Next JE
   Next IE
   DetalleComp = Ninguno
   Trans_No = 100
   SQL1 = "DELETE " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
   Ejecutar_SQL_SP SQL1
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
   TotalCajaMN = 0
   Total_Cheque = 0
   TotalIngreso = 0
   Fecha_Vence = MBFechaF
  'Recolectamos informacion
   Ln_No = 0
   NoCheque = Ninguno
   Trans_No = 100
   Ln_No = 0
   For IE = 0 To CantCtas - 1
      If CtasPro(IE).Cta <> "0" Then
         If CtasPro(IE).Valor >= 0 Then InsertarAsientos AdoAsiento, CtasPro(IE).Cta, 0, CtasPro(IE).Valor, 0
      End If
   Next IE
  'Ahora insertamos el Decimo respectivo por empleado
   Cta = SinEspaciosIzq(DCBanco)
   With AdoClientes.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           DetalleComp = .Fields("Cliente")
           CodigoCli = .Fields("Codigo")
           Select Case Decimo
             Case 3: Total = .Fields("Valor_Dec_3ro")
                     If .Fields("FormaPago10to") = "A" Then
                         InsertarAsientos AdoAsiento, Cta, 0, 0, Total
                     Else
                         InsertarAsientos AdoAsiento, Cta_CajaG, 0, 0, Total
                     End If
             Case 4: Total = .Fields("Valor_Dec_4to")
                     If .Fields("FormaPago10to") = "A" Then
                         InsertarAsientos AdoAsiento, Cta, 0, 0, Total
                     Else
                         InsertarAsientos AdoAsiento, Cta_CajaG, 0, 0, Total
                     End If
           End Select
          'MsgBox Total
         .MoveNext
       Loop
    End If
   End With

   RatonReloj
   Contador = 0
  'Asignamos Codigo Contable segun el Abono
   CodigoCli = Ninguno
   SumaDebe = 0: SumaHaber = 0
   SQL2 = "SELECT * " _
        & "FROM Asiento " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY A_No "
   Select_Adodc_Grid DGAsiento, AdoAsiento, SQL2
   With AdoAsiento.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           SumaDebe = SumaDebe + .Fields("DEBE")
           SumaHaber = SumaHaber + .Fields("HABER")
          .MoveNext
        Loop
    End If
   End With
   LabelIngresos.Caption = Format(SumaDebe, "#,##0.00")
   LabelEgresos.Caption = Format(SumaHaber, "#,##0.00")
   MsgBox "SE GENERO EL SIGUIENTE ARCHIVO: " & vbCrLf & vbCrLf & vbCrLf _
          & RutaGeneraFileAlumnos & vbCrLf & vbCrLf

End Sub

Public Sub InsValorCtaPro(NCta As String, NValor As Currency)
  For IE = 0 To CantCtas - 1
      If CtasPro(IE).Cta = NCta Then
         CtasPro(IE).Valor = CtasPro(IE).Valor + Redondear(NValor, 2)
      End If
  Next IE
End Sub

