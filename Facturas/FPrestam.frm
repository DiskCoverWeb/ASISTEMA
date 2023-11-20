VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FPrestamos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "FPrestam.frx":0000
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1365
      TabIndex        =   3
      Top             =   420
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCCodeudor 
      Bindings        =   "FPrestam.frx":001A
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1785
      TabIndex        =   9
      Top             =   1155
      Visible         =   0   'False
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCConyuge 
      Bindings        =   "FPrestam.frx":0034
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   1785
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin VB.CheckBox Check2 
      Caption         =   "CO-DEUDOR:"
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
      Top             =   1155
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CONYUGE:"
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
      Top             =   840
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir Pagaré"
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
      Left            =   1575
      Picture         =   "FPrestam.frx":004E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6510
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid DGPrestamo 
      Bindings        =   "FPrestam.frx":0918
      Height          =   3690
      Left            =   105
      TabIndex        =   28
      Top             =   2310
      Visible         =   0   'False
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   6509
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
   Begin MSDataListLib.DataCombo DCTipoPrestamo 
      Bindings        =   "FPrestam.frx":0932
      DataSource      =   "AdoTipoPrestamo"
      Height          =   315
      Left            =   105
      TabIndex        =   11
      Top             =   1890
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc AdoTipoPrestamo 
      Height          =   330
      Left            =   315
      Top             =   3570
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "TipoPrestamo"
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
      Left            =   315
      Top             =   3255
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.CommandButton Command4 
      Caption         =   "Tala &Cliente"
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
      Left            =   3570
      Picture         =   "FPrestam.frx":0950
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6510
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Prestamo"
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
      Left            =   5565
      Picture         =   "FPrestam.frx":11D2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6510
      Width           =   1905
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
      Height          =   855
      Left            =   7560
      Picture         =   "FPrestam.frx":1614
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6510
      Width           =   1905
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
      Left            =   7770
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   1890
      Width           =   1695
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
      Left            =   6930
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1890
      Width           =   750
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
      Left            =   6090
      MaxLength       =   49
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1890
      Width           =   750
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   3885
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSAdodcLib.Adodc AdoPrestamo 
      Height          =   330
      Left            =   315
      Top             =   2940
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Prestamo"
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
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "FPrestam.frx":200A
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   7560
      TabIndex        =   5
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   315
      Top             =   2625
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Autos"
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
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento No."
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
      TabIndex        =   4
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL CLIENTE"
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
      TabIndex        =   2
      Top             =   105
      Width           =   6210
   End
   Begin VB.Label LabelAbono 
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
      Left            =   7770
      TabIndex        =   22
      Top             =   6090
      Width           =   1695
   End
   Begin VB.Label LabelInteres 
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
      Left            =   4620
      TabIndex        =   27
      Top             =   6090
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Abono"
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
      TabIndex        =   23
      Top             =   6090
      Width           =   1275
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Interes"
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
      Left            =   3360
      TabIndex        =   26
      Top             =   6090
      Width           =   1275
   End
   Begin VB.Label LabelCapital 
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
      Left            =   1365
      TabIndex        =   25
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Capital"
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
      TabIndex        =   24
      Top             =   6090
      Width           =   1275
   End
   Begin VB.Label Label4 
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
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto del Crédito"
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
      TabIndex        =   16
      Top             =   1575
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
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
      TabIndex        =   14
      Top             =   1575
      Width           =   750
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
      Left            =   6090
      TabIndex        =   12
      Top             =   1575
      Width           =   750
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE CREDITO"
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
      Top             =   1575
      Width           =   5895
   End
End
Attribute VB_Name = "FPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
  If Check1.Value = 1 Then DCConyuge.Visible = True Else DCConyuge.Visible = False
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then DCCodeudor.Visible = True Else DCCodeudor.Visible = False
End Sub

Private Sub Command1_Click()
  RatonNormal
  Unload FPrestamos
End Sub

Private Sub Command2_Click()
'''  NumComp = ReadSetDataNum("Prestamos", True, False)
'''  Credito_No = Format(NumEmpresa, "000") & Format(NumComp, "000000")
'''  Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
'''           & "El Crédito " & Credito_No
'''  Titulo = "Formulario de Grabacion"
'''  If BoxMensaje = vbYes Then
'''     DGPrestamo.Visible = False
'''     NumComp = ReadSetDataNum("Prestamos", True, True)
'''     Credito_No = Format(NumEmpresa, "000") & Format(NumComp, "000000")
'''     CodigoP = SinEspaciosIzq(DCTipoPrestamo)
'''     If Convertir_Numero(TextContado, 2) > 0 Then
'''        Codigos = Format(Year(MBFecha), "0000")
'''        SetAdoAddNew "Clientes_Facturacion"
'''        SetAdoFields "T", Normal
'''        SetAdoFields "Codigo", CodigoCliente
'''        SetAdoFields "Valor", Convertir_Numero(TextContado, 2)
'''        SetAdoFields "Codigo_Inv", CodigoP
'''        SetAdoFields "Num_Mes", I
'''        SetAdoFields "Mes", MesesLetras(CInt(I))
'''        SetAdoFields "GrupoNo", Grupo_No
'''        SetAdoFields "Credito_No", Credito_No
'''        SetAdoFields "Fecha", MBFecha
'''        SetAdoFields "Codigo_Auto", SinEspaciosDer(DCAutos)
'''        SetAdoFields "Item", NumEmpresa
'''        SetAdoFields "Periodo", Codigos
'''        SetAdoFields "CodigoU", CodigoUsuario
'''        SetAdoUpdate
'''     End If
'''     sSQL = "SELECT TP,Fecha,Capital,Interes,Pagos,Saldo,Cuotas " _
'''          & "FROM Asiento_P " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " " _
'''          & "AND TP = '" & TipoFactura & "' "
'''     SelectDataGrid DGPrestamo, AdoPrestamo, sSQL
'''     With AdoPrestamo.Recordset
'''      If .RecordCount > 0 Then
'''         .MoveFirst
'''          Do While Not .EOF
'''             If .Fields("Cuotas") > 0 Then
'''                 I = Month(.Fields("Fecha"))
'''                 Valor = .Fields("Pagos")
'''                 Codigos = Format(Year(.Fields("Fecha")), "0000")
'''                 sSQL = "DELETE * " _
'''                      & "FROM Clientes_Facturacion " _
'''                      & "WHERE Codigo_Inv = '" & CodigoP & "' " _
'''                      & "AND Codigo = '" & CodigoCliente & "' " _
'''                      & "AND Periodo = '" & Codigos & "' " _
'''                      & "AND Num_Mes = " & CByte(I) & " " _
'''                      & "AND Item = '" & NumEmpresa & "' "
'''                 ConectarAdoExecute sSQL
'''                'Lineas de Creditos
'''                 SetAdoAddNew "Clientes_Facturacion"
'''                 SetAdoFields "T", Normal
'''                 SetAdoFields "Codigo", CodigoCliente
'''                 SetAdoFields "Valor", Valor
'''                 SetAdoFields "Codigo_Inv", CodigoP
'''                 SetAdoFields "Num_Mes", I
'''                 SetAdoFields "Mes", MesesLetras(CInt(I))
'''                 SetAdoFields "GrupoNo", Grupo_No
'''                 SetAdoFields "Credito_No", Credito_No
'''                 SetAdoFields "Fecha", .Fields("Fecha")
'''                 SetAdoFields "Item", NumEmpresa
'''                 SetAdoFields "Periodo", Codigos
'''                 SetAdoFields "CodigoU", CodigoUsuario
'''                 SetAdoUpdate
'''             End If
'''            .MoveNext
'''          Loop
'''         .MoveFirst
'''      End If
'''     End With
'''     DGPrestamo.Visible = True
'''     NumComp = ReadSetDataNum("Prestamos", True, False)
'''     FPrestamos.Caption = "PRESTAMO A PAZO FIJO."
'''     DGPrestamo.Caption = "Préstamo No. " & Format(NumEmpresa, "000") & Format(NumComp, "000000")
'''     sSQL = "DELETE * " _
'''          & "FROM Asiento_P " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " "
'''     ConectarAdoExecute sSQL
'''     sSQL = "SELECT TP,Fecha,Capital,Interes,Pagos,Saldo,Cuotas " _
'''          & "FROM Asiento_P " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " " _
'''          & "AND TP = '" & TipoFactura & "' "
'''     SelectDataGrid DGPrestamo, AdoPrestamo, sSQL
'''     MsgBox "Credito procesado con exito"
'''     MBFecha.SetFocus
''' End If
''' RatonNormal
End Sub

Private Sub Command3_Click()
'''  Cod_Bodega = ""
'''  Unidad = ""
'''  FechaValida MBFecha
'''  If Check1.Value = 1 Then Cod_Bodega = DCConyuge
'''  If Check2.Value = 1 Then Unidad = DCCodeudor
'''  Beneficiario = LblDescripcion.Caption
'''  Mifecha = MBFecha
'''  Factura_No = ReadSetDataNum("Prestamos", True, False)
'''  NumComp = ReadSetDataNum("Prestamos", True, False)
'''  Codigo2 = Format(NumEmpresa, "000") & Format(NumComp, "0000000")
'''  Mensajes = "Imprimir Acta Original"
'''  Titulo = "IMPRESION DE ACTAS"
'''  Producto = ""
'''  With AdoPrestamo.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''       Total = .Fields("Capital") + .Fields("Interes")
'''       Do While Not .EOF
'''          Producto = Producto & vbTab & vbTab & Format(.Fields("Cuotas"), "00") _
'''                   & vbTab & " " & .Fields("Fecha") & " " _
'''                   & " " & SetearBlancos(.Fields("Capital"), 12, 0, True, False, True) _
'''                   & " " & SetearBlancos(.Fields("Interes"), 12, 0, True, False, True) _
'''                   & " " & SetearBlancos(.Fields("Pagos"), 12, 0, True, False, True) _
'''                   & " " & SetearBlancos(.Fields("Saldo"), 12, 0, True, False, True) & " " & vbCrLf
'''          NoDias = .Fields("Cuotas")
'''         .MoveNext
'''       Loop
'''      .MoveFirst
'''   End If
'''  End With
'''  Cadena1 = "CONTRATO PAGARÉ"
'''
''' '[a]: ElNoAnio
''' '[B]: Beneficiario
''' '[d]: ElNoDias *
''' '[D]: DirCliente
''' '[E]: Empresa *
''' '[f]: MesesLetras(NoMeses)
''' '[F]: FechaStrgCiudad(Mifecha) *
''' '[p]: Interes *
''' '[P]: NoDias *
''' '[c]: NombreCiudad
''' '[C]: NombreCliente *
''' '[I]: CICliente *
''' '[m]: Producto
''' '[M]: LineaDeMalla
''' '[n]: ElValorUnit
''' '[t]: ElTotal (En letras) *
''' '[v]: ElValor
''' '[K]: Cod_Bodega
''' '[k]: Unidad
''' '[S]: Factura_No
''' '[A]: Carta_Porte
''' '[N]: ValorUnit
''' '[L]: LineasDeTexto
''' '[y]: Codigo2
''' '[T]: Total *
''' '[V]: Valor
'''  Imprimir_Documentos "PagareAutos.TXT", 8, 1.4, 1, 18, Cadena1, ""
'''  RatonNormal
End Sub

Private Sub Command4_Click()
  FechaValida MBFecha
  Mifecha = MBFecha
  Factura_No = ReadSetDataNum("Prestamos", True, False)
  NumComp = ReadSetDataNum("Prestamos", True, False)
  Codigo2 = Format(NumEmpresa, "000") & Format(NumComp, "0000000")
  Mensajes = "Imprimir Acta Original"
  Titulo = "IMPRESION DE ACTAS"
  Producto = ""
  With AdoPrestamo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Total = .Fields("Capital") + .Fields("Interes")
       Unidad = Cambio_Letras(Total)
       Do While Not .EOF
          Producto = Producto & vbTab & vbTab & Format(.Fields("Cuotas"), "00") _
                   & vbTab & " " & .Fields("Fecha") & " " _
                   & " " & SetearBlancos(.Fields("Capital"), 12, 0, True, False, True) _
                   & " " & SetearBlancos(.Fields("Interes"), 12, 0, True, False, True) _
                   & " " & SetearBlancos(.Fields("Pagos"), 12, 0, True, False, True) _
                   & " " & SetearBlancos(.Fields("Saldo"), 12, 0, True, False, True) & vbCrLf
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  Cadena1 = "TABLA DEL CLIENTE"
  Imprimir_Documentos "TablaCliente.TXT", 8, 1.4, 1, 18, Cadena1, ""
  RatonNormal
End Sub

Private Sub DCAutos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

'''Private Sub DCAutos_LostFocus()
'''  With AdoAutos.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''      .Find ("Codigo_Inv Like '" & SinEspaciosDer(DCAutos) & "' ")
'''       If Not .EOF Then
'''          LblDescripcion.Caption = .Fields("Producto") _
'''                         & ", Color: " & .Fields("Unidad") _
'''                         & ", Placa: " & .Fields("Codigo_Barra") _
'''                         & ", " & .Fields("Detalle") _
'''                         & ", Chasis: " & .Fields("Desc_Item") _
'''                         & ", Kilometraje: " & Format(.Fields("Maximo"), "#,##0")
'''       Else
'''          LblDescripcion.Caption = ""
'''       End If
'''   Else
'''       MsgBox "No existen datos"
'''       LblDescripcion.Caption = ""
'''   End If
'''  End With
'''
'''End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCClientes_LostFocus()
  DirCliente = "."
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCClientes.Text & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          NombreCliente = DCClientes.Text
          DireccionCli = .Fields("Direccion")
          TelefCliente = .Fields("Telefono")
          CICliente = .Fields("CI_RUC")
          DirCliente = DireccionCli
          Grupo_No = .Fields("Grupo")
       Else
          NombreCliente = DCClientes.Text
          MsgBox "Cliente no Asignado"
          DCClientes.SetFocus
       End If
   Else
       MsgBox "No existen datos"
       NombreCliente = DCClientes.Text
       MsgBox "Cliente no Asignado"
   End If
  End With
End Sub

Private Sub DCCodeudor_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCConyuge_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoPrestamo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  Trans_No = 200
  NumComp = ReadSetDataNum("Prestamos", True, False)
  Credito_No = Format(NumEmpresa, "000") & Format(NumComp, "000000")
  FPrestamos.Caption = "PRESTAMO A PAZO FIJO."
  DGPrestamo.Caption = "Préstamo No. " & Credito_No
  NuevoDiario = False
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  SelectDataGrid DGPrestamo, AdoPrestamo, sSQL
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Prestamo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CTP = 'C/D' "
  SelectDBCombo DCTipoPrestamo, AdoTipoPrestamo, sSQL, "Descripcion"
  
  sSQL = "SELECT (Producto & ' - ' & Codigo_Inv) As Productos,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND LEN(Codigo_Inv) > 1 " _
       & "ORDER BY Producto "
''''  SelectDBCombo DCAutos, AdoAutos, sSQL, "Productos"
  
'  NumComp = ReadSetDataNum("Prestamos", True, False)
  FPrestamos.Caption = "PRESTAMO A PAZO FIJO."
  DGPrestamo.Caption = "Préstamo No. " & Format(NumEmpresa, "000") & Format(NumComp, "000000")
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(False) & " " _
       & "ORDER BY Cliente "
  SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
  SelectDBCombo DCConyuge, AdoClientes, sSQL, "Cliente"
  SelectDBCombo DCCodeudor, AdoClientes, sSQL, "Cliente"
  RatonNormal
  MBFecha.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FPrestamos
  ConectarAdodc AdoPrestamo
  'ConectarAdodc AdoConsultar
  ConectarAdodc AdoTipoPrestamo
  ConectarAdodc AdoClientes
  ConectarAdodc AdoFactura
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

''''Private Sub TextContado_GotFocus()
''''  MarcarTexto TextContado
''''End Sub

Private Sub TextContado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInt_GotFocus()
  MarcarTexto TextInt
End Sub

Private Sub TextInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMeses_GotFocus()
  MarcarTexto TextMeses
End Sub

Private Sub TextMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextMonto_GotFocus()
  MarcarTexto TextMonto
End Sub

Private Sub TextMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
  
Private Sub TextMonto_LostFocus()
  FechaValida MBFecha
  TipoFactura = "FA"
  sSQL = "DELETE * " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND TP = '" & TipoFactura & "' "
  ConectarAdoExecute sSQL
'  IniciarAsientosDe DGAsientos, AdoAsientos
 'Monto del Prestamo
  TotalCapital = Convertir_Numero(TextMonto, 2)
  If TotalCapital > 0 Then
  
  Interes = Convertir_Numero(TextInt, 2)
  NoMeses = Convertir_Numero(TextMeses)
  If NoMeses <= 0 Then NoMeses = 1
 'Cuotas Iniciales
  NoCert = Convertir_Numero(TextNoCuotas)
  TotalAbonos = Convertir_Numero(TextValorCuotas, 2)
  Diferencia = NoCert * TotalAbonos
  TotalCredito = TotalCapital + ((TotalCapital * Interes) / 100)
  Abono = TotalCredito / NoMeses
 'Calculamos como seran las cuotas
  TotalCredito = Diferencia + TotalCredito
  'MsgBox Abono
  J = 0
  Mifecha = MBFecha
  SetAdoAddNew "Asiento_P"
  SetAdoFields "T_No", Trans_No
  SetAdoFields "TP", TipoFactura
  SetAdoFields "Fecha", Mifecha
  SetAdoFields "Capital", TotalCredito
  SetAdoFields "Interes", (Interes * TotalCapital) / 100
  SetAdoFields "Saldo", TotalCredito
  SetAdoFields "Cuotas", J
  SetAdoUpdate
  Saldo = TotalCredito
  'MsgBox Saldo
  J = J + 1
  If OpcInicio.Value And NoCert > 0 Then
     For I = 1 To NoCert
         Saldo = Saldo - TotalAbonos
         Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
         SetAdoAddNew "Asiento_P"
         SetAdoFields "T_No", Trans_No
         SetAdoFields "TP", TipoFactura
         SetAdoFields "Fecha", Mifecha
         SetAdoFields "Pagos", TotalAbonos
         SetAdoFields "Saldo", Saldo
         SetAdoFields "Cuotas", J
         SetAdoUpdate
         J = J + 1
     Next I
  End If
  'MsgBox NoMeses - NoCert
  
  For I = 1 To NoMeses
      Saldo = Saldo - Abono
      Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
      SetAdoAddNew "Asiento_P"
      SetAdoFields "T_No", Trans_No
      SetAdoFields "TP", TipoFactura
      SetAdoFields "Fecha", Mifecha
      SetAdoFields "Pagos", Abono
      SetAdoFields "Saldo", Saldo
      SetAdoFields "Cuotas", J
      SetAdoUpdate
      J = J + 1
  Next I
  If OpcFinal.Value And NoCert > 0 Then
     For I = 1 To NoCert
         Saldo = Saldo - TotalAbonos
         Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
         SetAdoAddNew "Asiento_P"
         SetAdoFields "T_No", Trans_No
         SetAdoFields "TP", TipoFactura
         SetAdoFields "Fecha", Mifecha
         SetAdoFields "Pagos", TotalAbonos
         SetAdoFields "Saldo", Saldo
         SetAdoFields "Cuotas", J
         SetAdoUpdate
         J = J + 1
     Next I
  End If
  TotalInteres = 0
  TotalCapital = 0
  DGPrestamo.Visible = False
  sSQL = "SELECT TP,Fecha,Capital,Interes,Pagos,Saldo,Cuotas " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND TP = '" & TipoFactura & "' "
  SelectDataGrid DGPrestamo, AdoPrestamo, sSQL
  With AdoPrestamo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          TotalCapital = TotalCapital + .Fields("Pagos")
          TotalInteres = TotalInteres + .Fields("Interes")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  DGPrestamo.Visible = True
  Codigo = Ninguno
  Cta_Interes = Ninguno
  Cta_Capital = Ninguno
''  With AdoTipoPrestamo.Recordset
''   If .RecordCount > 0 Then
''      .Find ("Descripcion = '" & DCTipoPrestamo.Text & "' ")
''       If Not .EOF Then
''          Cta_Interes = .Fields("Cta_Capital")
''          Cta_Capital = .Fields("Cta_Interes")
''          Codigo = .Fields("CTP")
''       End If
''  End If
''  End With
  'InsertarAsientos AdoAsientos, Cta_Capital, 0, TotalCapital, 0
  'InsertarAsientos AdoAsientos, Cta_Interes, 0, TotalInteres, 0
  'InsertarAsientos AdoAsientos, Cta_Ventas, 0, 0, TotalCapital + TotalInteres
  End If
  LabelCapital.Caption = Format(TotalCapital, "#,##0.00")
  LabelInteres.Caption = Format(TotalInteres, "#,##0.00")
  'LabelIVA.Caption = Format(Total_Comision, "#,##0.00")
  LabelAbono.Caption = Format(Total_Abonos, "#,##0.00")
End Sub

Private Sub TextNoCuotas_GotFocus()
  MarcarTexto TextNoCuotas
End Sub

Private Sub TextNoCuotas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValorCuotas_GotFocus()
  MarcarTexto TextValorCuotas
End Sub

Private Sub TextValorCuotas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
