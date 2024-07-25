VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Abonos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   8940
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   10185
      Picture         =   "Abonos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1050
      Width           =   1065
   End
   Begin VB.TextBox TextRecibido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   7980
      Width           =   1905
   End
   Begin VB.TextBox TextPorc 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5670
      MaxLength       =   10
      TabIndex        =   36
      Text            =   "0"
      Top             =   4830
      Width           =   435
   End
   Begin VB.TextBox TextRet 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   38
      Text            =   "0.00"
      Top             =   4620
      Width           =   1905
   End
   Begin VB.TextBox TextRetIVAS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   3885
      Width           =   1905
   End
   Begin VB.ComboBox CServicio 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5355
      TabIndex        =   30
      Text            =   "100"
      Top             =   4095
      Width           =   750
   End
   Begin VB.ComboBox CBienes 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5355
      TabIndex        =   25
      Text            =   "100"
      Top             =   3360
      Width           =   750
   End
   Begin VB.TextBox TextRetIVAB 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   3150
      Width           =   1905
   End
   Begin VB.TextBox TextBanco 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   42
      Top             =   5565
      Width           =   2325
   End
   Begin VB.TextBox TextCompRet 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      MaxLength       =   9
      TabIndex        =   20
      Text            =   "000000000"
      Top             =   2625
      Width           =   1170
   End
   Begin VB.TextBox TxtAutoRet 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2205
      MaxLength       =   49
      TabIndex        =   22
      Text            =   "0000000000"
      Top             =   2625
      Width           =   7890
   End
   Begin VB.TextBox TxtSerieRet 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "001001"
      Top             =   2625
      Width           =   855
   End
   Begin VB.TextBox TextCheqNo 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4095
      MaxLength       =   8
      TabIndex        =   41
      Top             =   5565
      Width           =   960
   End
   Begin VB.TextBox TextInteres 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   48
      Text            =   "0"
      Top             =   6300
      Width           =   2325
   End
   Begin VB.TextBox TextTotalBaucher 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   50
      Text            =   "0.00"
      Top             =   6090
      Width           =   1905
   End
   Begin VB.TextBox TextBaucher 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4095
      MaxLength       =   8
      TabIndex        =   47
      Top             =   6300
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCTarjeta 
      Bindings        =   "Abonos.frx":08CA
      DataSource      =   "AdoTarjeta"
      Height          =   360
      Left            =   105
      TabIndex        =   46
      Top             =   6300
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Banco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextCheque 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   14
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   5355
      Width           =   1905
   End
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "Abonos.frx":08E3
      DataSource      =   "AdoBanco"
      Height          =   360
      Left            =   105
      TabIndex        =   40
      Top             =   5565
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Banco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextCajaME 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   7140
      Width           =   1905
   End
   Begin VB.TextBox TextCajaMN 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8190
      MaxLength       =   14
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   6720
      Width           =   1905
   End
   Begin VB.TextBox TxtRecibo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2520
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0000000000"
      Top             =   105
      Width           =   1275
   End
   Begin VB.CheckBox CheqRecibo 
      Caption         =   "&INGRESO CAJA No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "Abonos.frx":08FA
      DataSource      =   "AdoFactura"
      Height          =   420
      Left            =   5250
      TabIndex        =   11
      Top             =   630
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "0000000000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   10185
      Top             =   4095
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   10185
      Top             =   3360
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
      Caption         =   "IngCaja"
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
      Left            =   10185
      Top             =   5145
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
   Begin MSAdodcLib.Adodc AdoRecibo 
      Height          =   330
      Left            =   10185
      Top             =   3045
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
      Caption         =   "Recibo"
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   10185
      Picture         =   "Abonos.frx":0913
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   435
      Left            =   8715
      TabIndex        =   5
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   10185
      Top             =   4410
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   10185
      Top             =   5460
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
      Caption         =   "DetAcomp"
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "Abonos.frx":0D55
      DataSource      =   "AdoDetAcomp"
      Height          =   420
      Left            =   2100
      TabIndex        =   7
      Top             =   630
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "FA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTarjeta 
      Height          =   330
      Left            =   10185
      Top             =   1995
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
      Caption         =   "Tarjeta"
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
   Begin MSDataListLib.DataCombo DCRetIBienes 
      Bindings        =   "Abonos.frx":0D6F
      DataSource      =   "AdoRetIvaBienes"
      Height          =   360
      Left            =   105
      TabIndex        =   24
      Top             =   3360
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCRetISer 
      Bindings        =   "Abonos.frx":0D8D
      DataSource      =   "AdoRetIvaServicio"
      Height          =   360
      Left            =   105
      TabIndex        =   29
      Top             =   4095
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCRetFuente 
      Bindings        =   "Abonos.frx":0DAD
      DataSource      =   "AdoRetFuente"
      Height          =   360
      Left            =   105
      TabIndex        =   34
      Top             =   4830
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCodRet 
      Bindings        =   "Abonos.frx":0DC8
      DataSource      =   "AdoCodRet"
      Height          =   360
      Left            =   4620
      TabIndex        =   35
      ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
      Top             =   4830
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   10185
      Top             =   2310
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
      Caption         =   "RetIvaBienes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicio 
      Height          =   330
      Left            =   10185
      Top             =   4725
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
      Caption         =   "RetIvaServicio"
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   10185
      Top             =   5775
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
      Caption         =   "RetFuente"
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
   Begin MSAdodcLib.Adodc AdoCodRet 
      Height          =   330
      Left            =   10185
      Top             =   3675
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
      Caption         =   "CodRet"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "Abonos.frx":0DE0
      DataSource      =   "AdoSerie"
      Height          =   420
      Left            =   3570
      TabIndex        =   9
      Top             =   630
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "001001"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "Abonos.frx":0DF7
      DataSource      =   "AdoAutorizacion"
      Height          =   420
      Left            =   1470
      TabIndex        =   15
      Top             =   1155
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   10185
      Top             =   6195
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Serie"
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
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   10185
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
      Caption         =   "Autorizacion"
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
   Begin MSDataListLib.DataCombo DCVendedor 
      Bindings        =   "Abonos.frx":0E15
      DataSource      =   "AdoRecibo"
      Height          =   330
      Left            =   105
      TabIndex        =   64
      Top             =   8505
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VENDEDOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   63
      Top             =   8190
      Width           =   4845
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   105
      X2              =   10080
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   2940
      TabIndex        =   8
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   105
      TabIndex        =   14
      Top             =   1155
      Width           =   1380
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUENTA DEL BANCO                                                    CHEQUE      NOMBRE DEL BANCO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   39
      Top             =   5250
      Width           =   7260
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RETENCION EN LA FUENTE                                                    CODIGO           %"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   33
      Top             =   4515
      Width           =   6000
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RETENCION DEL I.V.A. EN SERVICIO                                                      %"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   28
      Top             =   3780
      Width           =   6000
   End
   Begin VB.Label LabelCambio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   8190
      TabIndex        =   62
      Top             =   8400
      Width           =   1905
   End
   Begin VB.Label LabelPend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   8190
      TabIndex        =   58
      Top             =   7560
      Width           =   1905
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAMBIO A ENTREGAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   61
      Top             =   8400
      Width           =   3165
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR RECIBIDO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   59
      Top             =   7980
      Width           =   3165
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO ACTUAL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   57
      Top             =   7560
      Width           =   3165
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Retenido"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6195
      TabIndex        =   37
      Top             =   4620
      Width           =   2010
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Retenido"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6195
      TabIndex        =   31
      Top             =   3885
      Width           =   2010
   End
   Begin VB.Label Label24 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RETENCION DEL I.V.A. EN BIENES                                                          %"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   23
      Top             =   3045
      Width           =   6000
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor Retenido"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6195
      TabIndex        =   26
      Top             =   3150
      Width           =   2010
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Retención &No. "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   18
      Top             =   2310
      Width           =   2010
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorizacion"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   21
      Top             =   2310
      Width           =   7890
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7455
      TabIndex        =   43
      Top             =   5355
      Width           =   750
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7455
      TabIndex        =   49
      Top             =   6090
      Width           =   750
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TARJETA DE CREDITO                                                 BAUCHER    INTERES DE LA TARJETA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   45
      Top             =   5985
      Width           =   7260
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja ME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   55
      Top             =   7140
      Width           =   3165
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Caja MN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   53
      Top             =   6720
      Width           =   3165
   End
   Begin VB.Label LabelDolares 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5355
      TabIndex        =   3
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COTIZACION"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3885
      TabIndex        =   2
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label LblNota 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   645
      Left            =   105
      TabIndex        =   52
      Top             =   7455
      Width           =   4845
   End
   Begin VB.Label LblObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   645
      Left            =   105
      TabIndex        =   51
      Top             =   6720
      Width           =   4845
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   8190
      TabIndex        =   13
      Top             =   630
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6825
      TabIndex        =   12
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo de Documento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   105
      TabIndex        =   6
      Top             =   630
      Width           =   2010
   End
   Begin VB.Label LblGrupo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Grupo No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   8610
      TabIndex        =   17
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   4725
      TabIndex        =   10
      Top             =   630
      Width           =   540
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA DEL ABONO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6615
      TabIndex        =   4
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label LblCliente 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CLIENTE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   105
      TabIndex        =   16
      Top             =   1680
      Width           =   8520
   End
End
Attribute VB_Name = "Abonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EsNotaDeVenta As Boolean

Private Sub Command1_Click()
  If CFechaLong(MBFecha) < CFechaLong(FechaCorte) Then
     MsgBox "No se puede grabar abonos con fecha inferior a la emision de la factura"
     MBFecha.SetFocus
  Else
      TextoValido TextBanco
      TextoValido TextCheqNo
      TA.Recibi_de = NombreCliente
      Mensajes = "Esta Seguro que desea grabar estos pagos."
      Titulo = "Formulario de Grabación."
      If BoxMensaje = vbYes Then
         CodigoVen = CodigoUsuario
         If ComisionEjec Then
            With AdoRecibo.Recordset
             If .RecordCount > 0 Then
                .MoveFirst
                .Find ("Nombre_Completo = '" & DCVendedor & "' ")
                 If Not .EOF Then CodigoVen = .fields("Codigo")
             End If
            End With
         End If
         Calculo_Saldo
         FechaTexto = FechaSistema
         If CheqRecibo.value = 1 Then
            DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
         Else
            DiarioCaja = Val(TxtRecibo)
         End If
         TotalAbonos = 0
         Cta = SinEspaciosIzq(DCBanco.Text)
         Cta1 = SinEspaciosIzq(DCTarjeta.Text)
         NombreBanco1 = TrimStrg(MidStrg(DCTarjeta.Text, Len(Cta1) + 1, 60))
         NombreBanco1 = MidStrg(NombreBanco1, 1, 25)
         TA.Cta_CxP = Cta_Cobrar
         TA.CodigoC = CodigoCliente
         TA.T = Normal
         TA.TP = FA.TC
         TA.Serie = FA.Serie
         TA.Factura = FA.Factura
         TA.Fecha = MBFecha
         
        'Abono de Factura Caja MN
         TA.Cta = Cta_CajaG
         TA.Banco = "EFECTIVO MN"
         TA.Cheque = Grupo_No
         TA.Abono = TotalCajaMN
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos TA
         
        'Abono de Factura Caja ME
         TA.Cta = Cta_CajaGE
         TA.Banco = "EFECTIVO MN"
         TA.Cheque = Grupo_No
         TA.Abono = TotalCajaME
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos TA
         
        'Abono de Factura Banco
         TA.Cta = TrimStrg(SinEspaciosIzq(DCBanco))
         TA.Banco = TextBanco
         TA.Cheque = TextCheqNo
         TA.Abono = Total_Bancos
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos TA
         
        'Abono de Factura Tarjeta
         TA.Cta = TrimStrg(SinEspaciosIzq(DCTarjeta))
         TA.Banco = NombreBanco1
         TA.Cheque = TextBaucher
         TA.Abono = Total_Tarjeta
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos TA
         
        'Abono de Factura Rete. IVA Bienes
         Codigo1 = Format$(Val(TrimStrg(MidStrg(TxtSerieRet, 1, 3))), "000")
         Codigo2 = Format$(Val(TrimStrg(MidStrg(TxtSerieRet, 4, 3))), "000")
         
         TA.Cta = TrimStrg(SinEspaciosIzq(DCRetIBienes))
         TA.Banco = "RETENCION IVA BIENES"
         TA.Cheque = TextCompRet
         TA.Abono = Total_RetIVAB
         TA.AutorizacionR = TxtAutoRet
         TA.Establecimiento = Codigo1
         TA.Emision = Codigo2
         TA.Porcentaje = Val(CBienes)
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos_Retenciones TA
         
        'Abono de Factura Ret IVA Servicio
         TA.Cta = TrimStrg(SinEspaciosIzq(DCRetISer))
         TA.Banco = "RETENCION IVA SERVICIO"
         TA.Cheque = TextCompRet
         TA.Abono = Total_RetIVAS
         TA.AutorizacionR = TxtAutoRet
         TA.Establecimiento = Codigo1
         TA.Emision = Codigo2
         TA.Porcentaje = Val(CServicio)
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos_Retenciones TA
         
        'Abono de Factura Ret. Fuente
         TA.Cta = TrimStrg(SinEspaciosIzq(DCRetFuente))
         TA.Banco = "RETENCION FUENTE - " & DCCodRet
         TA.Cheque = TextCompRet
         TA.Abono = Total_Ret
         TA.AutorizacionR = TxtAutoRet
         TA.Establecimiento = Codigo1
         TA.Emision = Codigo2
         TA.Porcentaje = Val(TextPorc)
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos_Retenciones TA
         
        'Abono de Factura Interes Tarjeta
         TA.TP = "TJ"
         TA.Cta = Cta1
         TA.Cta_CxP = Cta_Tarjetas
         TA.Banco = "INTERES POR TARJETA"
         TA.Cheque = TextBaucher
         TA.Abono = Val(TextInteres)
         If TA.Abono > 0 Then TotalAbonos = TotalAbonos + TA.Abono
         Grabar_Abonos TA
         
         Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
         
    '''     T = "P"
    '''     If SaldoDisp <= 0 Then
    '''        T = "C"
    '''        SaldoDisp = 0
    '''     End If
    '''     sSQL = "UPDATE Facturas " _
    '''          & "SET Saldo_MN = " & SaldoDisp & ",T = '" & T & "' " _
    '''          & "WHERE Item = '" & NumEmpresa & "' " _
    '''          & "AND Periodo = '" & Periodo_Contable & "' " _
    '''          & "AND TC = '" & TA.TP & "' " _
    '''          & "AND Serie = '" & TA.Serie & "' " _
    '''          & "AND Autorizacion = '" & TA.Autorizacion & "' " _
    '''          & "AND Factura = " & TA.Factura & " " _
    '''          & "AND CodigoC = '" & TA.CodigoC & "' "
    '''     Ejecutar_SQL_SP sSQL
         
         LabelCambio.Caption = "0.00"
         TextCajaMN = "0.00"
         TextCajaME = "0.00"
         TextRet = "0.00"
         TextRetIVA = "0.00"
         TextCheqNo = ""
         TextCheque = "0.00"
         TextBaucher = ""
         TextTotalBaucher = "0.00"
         TextRecibido = "0.00"
         TextInteres = "0.00"
         RatonNormal
         Imprimir_Comprobante_Caja TA
      End If
      If Evaluar Then Unload Abonos
  End If
End Sub

Private Sub Command3_Click()
   Unload Abonos
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCodRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCodRet_LostFocus()
  TextPorc = "0"
  With AdoCodRet.Recordset
   If .RecordCount Then
      .MoveFirst
      .Find ("Codigo = '" & DCCodRet & "' ")
       If Not .EOF Then TextPorc = .fields("Porcentaje")
   End If
  End With
End Sub

Private Sub DCFactura_GotFocus()
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Total_IVA = 0
  Total_Ret = 0
  Total_RetIVAB = 0
  Total_RetIVAS = 0
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
   'If KeyCode = vbKeyEscape Then DCCliente.SetFocus
End Sub

Private Sub DCFactura_LostFocus()
 'Listamos las autorizaciones de facturas pendientes por cliente
  FA.Factura = Val(DCFactura)
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND Serie = '" & FA.Serie & "' " _
       & "AND Factura = " & FA.Factura & " "
  If Evaluar Then
     sSQL = sSQL & "AND Autorizacion = '" & FA.Autorizacion & "' "
  Else
'     sSQL = sSQL & "AND CodigoC = '" & CodigoCliente & "' "
  End If
  sSQL = sSQL & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
End Sub

Private Sub DCTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  FA.TC = DCTipo
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub Form_Activate()
  ComisionEjec = Leer_Campo_Empresa("Comision_Ejecutivo")
  Evaluar = True
  
  With FA
   If .Autorizacion = Ninguno And _
      .Serie = Ninguno And _
      .TC = Ninguno And _
      .Factura = 0 Then Evaluar = False
  End With
 ' MsgBox "Documento " & FA.TC & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000")
  ControlEsNumerico TextCajaMN
  ControlEsNumerico TextCajaME
  ControlEsNumerico TextCheque
  ControlEsNumerico TextTotalBaucher
  If FechaTexto = "" Or FechaTexto = Ninguno Or IsNumeric(FechaTexto) Then FechaTexto = FechaSistema
  MBFecha = FechaTexto
  CBienes.Clear
  CBienes.AddItem "0"
  CBienes.AddItem "10"
  CBienes.AddItem "30"
  CBienes.AddItem "100"
  CBienes.Text = "0"
  
  CServicio.Clear
  CServicio.AddItem "0"
  CServicio.AddItem "20"
  CServicio.AddItem "70"
  CServicio.AddItem "100"
  CServicio.Text = "0"
 
  If ComisionEjec Then
     sSQL = "SELECT RP.Codigo, A.Usuario, A.Nombre_Completo " _
          & "FROM Accesos As A, Catalogo_Rol_Pagos As RP " _
          & "WHERE RP.Item = '" & NumEmpresa & "' " _
          & "AND RP.Periodo = '" & Periodo_Contable & "' " _
          & "AND A.Codigo = RP.Codigo " _
          & "ORDER BY A.Nombre_Completo "
     SelectDB_Combo DCVendedor, AdoRecibo, sSQL, "Nombre_Completo"
     With AdoRecibo.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Codigo = '" & CodigoUsuario & "' ")
          If Not .EOF Then DCVendedor = .fields("Nombre_Completo")
      End If
     End With
  Else
     DCVendedor.Text = CodigoUsuario
     CodigoVen = CodigoUsuario
     DCVendedor.Visible = False
  End If
  
 'Procesar_Saldo_De_Facturas Abonos, AdoFactura
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format$(DiarioCaja, "0000000") Else TxtRecibo = ""
  'If Factura_No = 0 Then Command2.Caption = "&Salir"
  Mifecha = BuscarFecha(FechaSistema)
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC IN ('BA','CJ','CP','C','P') " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'TJ' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCTarjeta, AdoTarjeta, sSQL, "NomCuenta"
  
 'Carga los Conceptos de Retención en la Fuente
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CF' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
 'Carga los Conceptos de retención IVA Servicios al DataCombo
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CI' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetISer, AdoRetIvaServicio, sSQL, "Cuentas"
 'Carga los Conceptos de retención IVA Bienes al DataCombo
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CB' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetIBienes, AdoRetIvaBienes, sSQL, "Cuentas"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT * " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFecha) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBFecha) & "# " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCodRet, AdoCodRet, sSQL, "Codigo"
  
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND NOT TC IN ('OP','C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDB_Combo DCTipo, AdoDetAcomp, sSQL, "TC"
'  Listar_Clientes_F
  RatonNormal
'''  If Evaluar Then
'''     DCCliente.Text = NombreCliente
'''     DCCliente.SetFocus
'''  Else
'''    'TxtRecibo.SetFocus
'''  End If
End Sub

Private Sub Form_Load()
   CentrarForm Abonos
   ConectarAdodc AdoSerie
   ConectarAdodc AdoBanco
   ConectarAdodc AdoTarjeta
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoRecibo
   ConectarAdodc AdoCliente
   ConectarAdodc AdoAutorizacion
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoCodRet
   ConectarAdodc AdoRetFuente
   ConectarAdodc AdoRetIvaBienes
   ConectarAdodc AdoRetIvaServicio
   FechaTexto = FechaSistema
   FA.TC = Ninguno
   FA.Serie = Ninguno
   FA.Factura = 0
   Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha, True
 'Carga los conceptos de Retencion segun la fecha de Registro
  If Not IsDate(MBFecha) Then MBFecha = FechaSistema
  sSQL = "SELECT " & Full_Fields("Tipo_Concepto_Retencion") & " " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFecha) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBFecha) & "# " _
       & "ORDER BY Codigo "
  'MsgBox sSQL
  SelectDB_Combo DCCodRet, AdoCodRet, sSQL, "Codigo"
End Sub

Private Sub TextBanco_GotFocus()
  MarcarTexto TextBanco
End Sub

Private Sub TextBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextBanco_LostFocus()
  TextoValido TextBanco
End Sub

Private Sub TextBaucher_GotFocus()
  MarcarTexto TextBaucher
End Sub

Private Sub TextBaucher_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCajaME_GotFocus()
  MarcarTexto TextCajaME
End Sub

Private Sub TextCajaME_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCajaME_LostFocus()
  TextoValido TextCajaME, True
  TotalCajaME = Redondear(Val(CCur(TextCajaME.Text)), 2)
  TextCajaME.Text = Format$(TotalCajaME, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
End Sub

Private Sub TextCajaMN_GotFocus()
  TextCajaMN.Text = Saldo
  MarcarTexto TextCajaMN
End Sub

Private Sub TextCajaMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCajaMN_LostFocus()
  'TextCajaMN.Text = Format$(TextCajaMN.Text, "#,##0.00")
  TextoValido TextCajaMN, True
 'MsgBox TextCajaMN.Text
  TotalCajaMN = Redondear(Val(CCur(TextCajaMN.Text)), 2)
  'TextCajaMN.Text = Format$(TotalCajaMN, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextCheque_GotFocus()
  MarcarTexto TextCheque
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCheque_LostFocus()
  TextoValido TextCheque, True
  Total_Bancos = Redondear(Val(CCur(TextCheque.Text)), 2)
  TextCheque.Text = Format$(Total_Bancos, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextInteres_GotFocus()
  MarcarTexto TextInteres
End Sub

Private Sub TextInteres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInteres_LostFocus()
  If MidStrg(TextInteres, Len(TextInteres), 1) = "%" Then
     Valor = MidStrg(TextInteres, 1, Len(TextInteres) - 1)
     TextInteres = Valor * Val(LabelPend.Caption) / 100
  End If
  TextoValido TextInteres, True
End Sub

Private Sub TextPorc_GotFocus()
  MarcarTexto TextPorc
End Sub

Private Sub TextPorc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRecibido_GotFocus()
  Calculo_Saldo
  TextRecibido = Format$(TotalAbonos + Val(CCur(TextInteres)), "#,##0.00")
  MarcarTexto TextRecibido
End Sub

Private Sub TextRecibido_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRecibido_LostFocus()
  TextoValido TextRecibido, True
  Calculo_Saldo
  LabelCambio.Caption = Format$(Val(CCur(TextRecibido)) - TotalAbonos - Val(CCur(TextInteres)), "#,##0.00")
End Sub

Private Sub TextRet_GotFocus()
  MarcarTexto TextRet
End Sub

Private Sub TextRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRet_LostFocus()
  TextoValido TextRet, True
  Total_Ret = Redondear(Val(CCur(TextRet)), 2)
  TextRet = Format$(Total_Ret, "#,##0.00")
  Calculo_Saldo
End Sub

''Private Sub TextTarjeta_GotFocus()
''  MarcarTexto TextTarjeta
''End Sub

''Private Sub TextTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
''   PresionoEnter KeyCode
''End Sub

Private Sub TextTotalBaucher_GotFocus()
  MarcarTexto TextTotalBaucher
End Sub

Private Sub TextTotalBaucher_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextTotalBaucher_LostFocus()
  TextoValido TextTotalBaucher, True
  Total_Tarjeta = Redondear(Val(CCur(TextTotalBaucher.Text)), 2)
  TextTotalBaucher.Text = Format$(Total_Tarjeta, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Public Sub Calculo_Saldo()
   TotalAbonos = TotalCajaMN + TotalCajaME + Total_Bancos + Total_Tarjeta + Total_IVA + Total_Ret + Total_RetIVAB + Total_RetIVAS
   SaldoDisp = Saldo - TotalAbonos
End Sub

Public Sub Listar_Facturas_P()
    Select Case FA.TC
      Case "PV": Label2.Caption = " Punto de Venta No."
      Case "NV": Label2.Caption = " Nota de Venta No."
      Case Else: Label2.Caption = " Factura No."
    End Select

    SQL1 = "SELECT F.TC,F.Factura,F.Autorizacion,F.Serie,F.CodigoC,F.Fecha,F.Fecha_V," _
         & "F.Total_MN,F.Total_ME,F.Saldo_MN,F.Saldo_ME,F.Cta_CxP,F.Nota,F.Observacion,F.Cotizacion," _
         & "C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo " _
         & "FROM Facturas As F INNER JOIN Clientes As C " _
         & "ON F.CodigoC = C.Codigo " _
         & "WHERE F.T = '" & Pendiente & "' " _
         & "AND F.Item = '" & NumEmpresa & "' " _
         & "AND F.Periodo = '" & Periodo_Contable & "' " _
         & "AND F.Serie = '" & FA.Serie & "' " _
         & "AND F.TC = '" & FA.TC & "' " _
         & "AND F.Saldo_MN > 0 " _
         & "ORDER BY F.TC, F.Factura "
    SelectDB_Combo DCFactura, AdoFactura, SQL1, "Factura"
    If AdoFactura.Recordset.RecordCount <= 0 Then TxtRecibo.SetFocus
End Sub

'''Public Sub Listar_Clientes_F()
'''  Codigo1 = DCRecibo
'''  If Codigo1 = "" Then Codigo1 = Ninguno
'''  sSQL = "SELECT Grupo,Cliente,Codigo,CI_RUC,C.Direccion,C.Telefono,COUNT(F.Factura) As CantFacturas " _
'''       & "FROM Clientes As C,Facturas As F " _
'''       & "WHERE F.Item = '" & NumEmpresa & "' " _
'''       & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND F.TC = '" & FA.TC & "' " _
'''       & "AND F.Serie = '" & FA.Serie & "' " _
'''       & "AND F.T = 'P' " _
'''       & "AND F.TC IN ('PV','FA','NV') " _
'''       & "AND F.Saldo_MN > 0 " _
'''       & "AND C.Codigo = F.CodigoC "
'''  If Evaluar Then
'''     sSQL = sSQL _
'''          & "AND F.Factura = " & FA.Factura & " " _
'''          & "AND F.Autorizacion = '" & FA.Autorizacion & "' " _
'''          & "AND F.CodigoC = '" & FA.CodigoC & "' "
'''  Else
''''     If Codigo1 <> Ninguno Then sSQL = sSQL & "AND C.Grupo = '" & Codigo1 & "' "
'''  End If
'''  sSQL = sSQL & "GROUP BY Grupo,Cliente,Codigo,CI_RUC,C.Direccion,C.Telefono " _
'''       & "ORDER BY Cliente "
'''  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
'''  If AdoCliente.Recordset.RecordCount <= 0 Then CheqRecibo.SetFocus
'''End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format$(Val(TxtRecibo), "0000000")
End Sub

Private Sub TextRetIVAB_GotFocus()
  MarcarTexto TextRetIVAB
End Sub

Private Sub TextRetIVAB_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextRetIVAB_LostFocus()
  TextoValido TextRetIVAB, True
  Total_RetIVAB = Redondear(Val(CCur(TextRetIVAB)), 2)
  TextRetIVAB = Format$(Total_RetIVAB, "#,##0.00")
  Calculo_Saldo
End Sub

Private Sub TextRetIVAS_GotFocus()
  MarcarTexto TextRetIVAS
End Sub

Private Sub TextRetIVAS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextRetIVAS_LostFocus()
  TextoValido TextRetIVAS, True
  Total_RetIVAS = Redondear(Val(CCur(TextRetIVAS)), 2)
  TextRetIVAS = Format$(Total_RetIVAS, "#,##0.00")
  Calculo_Saldo
End Sub

Private Sub TxtSerieRet_GotFocus()
  MarcarTexto TxtSerieRet
End Sub

Private Sub TxtSerieRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtSerieRet_LostFocus()
  TxtSerieRet = Format$(Val(TxtSerieRet), "000000")
End Sub

Private Sub TxtAutoRet_GotFocus()
  MarcarTexto TxtAutoRet
End Sub

Private Sub TxtAutoRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtAutoRet_LostFocus()
  TextoValido TxtAutoRet
  If Len(TxtAutoRet) < 10 Then TxtAutoRet = Format$(Val(TxtAutoRet), "0000000000")
End Sub

Private Sub TextCompRet_GotFocus()
  MarcarTexto TextCompRet
End Sub

Private Sub TextCompRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCompRet_LostFocus()
  TextCompRet = Format$(Val(TextCompRet), "0000000")
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAutorizacion_LostFocus()
 'Listamos las autorizaciones de facturas pendientes por cliente
  FA.Autorizacion = DCAutorizacion
  Factura_No = 0: Saldo = 0: Cotizacion = 0: TotalDolar = 0
  Saldo_ME = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura Like '" & DCFactura.Text & "' ")
       If Not .EOF Then
          Factura_No = .fields("Factura")
          LblObs.Caption = " " & .fields("Observacion")
          LblNota.Caption = " " & .fields("Nota")
          Cta_Cobrar = .fields("Cta_CxP")
          CodigoCliente = .fields("CodigoC")
          Saldo = Redondear(.fields("Saldo_MN"), 2)
          Saldo_ME = Redondear(.fields("Saldo_ME"), 2)
          TipoFactura = .fields("TC")
          TotalDolar = .fields("Total_ME")
          Cotizacion = .fields("Cotizacion")
          TA.Serie = .fields("Serie")
          TA.Autorizacion = .fields("Autorizacion")
          TA.CI_RUC_Cli = .fields("CI_RUC")
          NombreCliente = .fields("Cliente")
          DireccionCli = .fields("Direccion")
          Grupo_No = .fields("Grupo")
          LblCliente.Caption = " " & NombreCliente
          LblGrupo.Caption = " " & Grupo_No
          
          FechaCorte = .fields("Fecha")
          TA.Fecha = MBFecha
          
          If TotalDolar <> 0 Then
             TextRet.Enabled = True
          Else
            TextCajaMN.Enabled = True
            TextCheque.Enabled = True
          End If
          Efectivo = 0: Cheque = 0: Retencion = 0
          LabelDolares.Caption = Format$(Cotizacion, "#,##0.00")
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
          If TotalDolar <> 0 Then
             TextRet.SetFocus
          Else
             TextCajaMN.SetFocus
          End If
          Abonos.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
          If CFechaLong(MBFecha) < CFechaLong(FechaCorte) Then
             MsgBox "No se puede grabar abonos con fecha inferior a la emision de la factura"
             MBFecha.SetFocus
          End If
       End If
    End If
  End With
  Calculo_Saldo
  LabelPend.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  FA.Serie = DCSerie
  Listar_Facturas_P
End Sub

