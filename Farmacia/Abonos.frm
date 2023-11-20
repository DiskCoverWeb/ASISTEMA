VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Abonos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   8205
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   11385
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
   ScaleHeight     =   8205
   ScaleWidth      =   11385
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
      TabIndex        =   62
      Text            =   "0.00"
      Top             =   7245
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
      TabIndex        =   40
      Text            =   "0"
      Top             =   4095
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
      TabIndex        =   42
      Text            =   "0.00"
      Top             =   3885
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
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   3150
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
      TabIndex        =   34
      Text            =   "100"
      Top             =   3360
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
      TabIndex        =   29
      Text            =   "100"
      Top             =   2625
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
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   2415
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
      TabIndex        =   46
      Top             =   4830
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
      Height          =   465
      Left            =   2415
      MaxLength       =   7
      TabIndex        =   22
      Text            =   "0000000"
      Top             =   1785
      Width           =   1065
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
      Height          =   465
      Left            =   4830
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "0000000000"
      Top             =   1785
      Width           =   1275
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
      Height          =   465
      Left            =   1575
      MaxLength       =   8
      TabIndex        =   21
      Text            =   "001001"
      Top             =   1785
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
      TabIndex        =   45
      Top             =   4830
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
      TabIndex        =   52
      Text            =   "0"
      Top             =   5565
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
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   5355
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
      TabIndex        =   51
      Top             =   5565
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCTarjeta 
      Bindings        =   "Abonos.frx":08CA
      DataSource      =   "AdoTarjeta"
      Height          =   360
      Left            =   105
      TabIndex        =   50
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
      TabIndex        =   48
      Text            =   "0.00"
      Top             =   4620
      Width           =   1905
   End
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "Abonos.frx":08E3
      DataSource      =   "AdoBanco"
      Height          =   360
      Left            =   105
      TabIndex        =   44
      Top             =   4830
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
      TabIndex        =   58
      Text            =   "0.00"
      Top             =   6405
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
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   5985
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
      Caption         =   "&RECIBO DE CAJA No."
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
      Left            =   8400
      TabIndex        =   17
      Top             =   1155
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCRecibo 
      Bindings        =   "Abonos.frx":0913
      DataSource      =   "AdoRecibo"
      Height          =   420
      Left            =   945
      TabIndex        =   7
      Top             =   630
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   741
      _Version        =   393216
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Abonos.frx":092B
      DataSource      =   "AdoCliente"
      Height          =   420
      Left            =   3570
      TabIndex        =   9
      Top             =   630
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   741
      _Version        =   393216
      Text            =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      Picture         =   "Abonos.frx":0944
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   435
      Left            =   8925
      TabIndex        =   5
      Top             =   105
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
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
      Bindings        =   "Abonos.frx":0D86
      DataSource      =   "AdoDetAcomp"
      Height          =   420
      Left            =   5670
      TabIndex        =   15
      Top             =   1155
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "FA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      Bindings        =   "Abonos.frx":0DA0
      DataSource      =   "AdoRetIvaBienes"
      Height          =   360
      Left            =   105
      TabIndex        =   28
      Top             =   2625
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
      Bindings        =   "Abonos.frx":0DBE
      DataSource      =   "AdoRetIvaServicio"
      Height          =   360
      Left            =   105
      TabIndex        =   33
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
   Begin MSDataListLib.DataCombo DCRetFuente 
      Bindings        =   "Abonos.frx":0DDE
      DataSource      =   "AdoRetFuente"
      Height          =   360
      Left            =   105
      TabIndex        =   38
      Top             =   4095
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
      Bindings        =   "Abonos.frx":0DF9
      DataSource      =   "AdoCodRet"
      Height          =   360
      Left            =   4620
      TabIndex        =   39
      ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
      Top             =   4095
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
      Bindings        =   "Abonos.frx":0E11
      DataSource      =   "AdoSerie"
      Height          =   420
      Left            =   3780
      TabIndex        =   13
      Top             =   1155
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "001001"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "Abonos.frx":0E28
      DataSource      =   "AdoAutorizacion"
      Height          =   420
      Left            =   1470
      TabIndex        =   11
      Top             =   1155
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
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
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   105
      X2              =   10080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   3150
      TabIndex        =   12
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   105
      TabIndex        =   10
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
      TabIndex        =   43
      Top             =   4515
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
      TabIndex        =   37
      Top             =   3780
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
      TabIndex        =   32
      Top             =   3045
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
      TabIndex        =   64
      Top             =   7665
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
      TabIndex        =   60
      Top             =   6825
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
      TabIndex        =   63
      Top             =   7665
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
      TabIndex        =   61
      Top             =   7245
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
      TabIndex        =   59
      Top             =   6825
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
      TabIndex        =   41
      Top             =   3885
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
      TabIndex        =   35
      Top             =   3150
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
      TabIndex        =   27
      Top             =   2310
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
      TabIndex        =   30
      Top             =   2415
      Width           =   2010
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retención No. "
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
      TabIndex        =   20
      Top             =   1785
      Width           =   1485
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
      Height          =   435
      Left            =   3465
      TabIndex        =   23
      Top             =   1785
      Width           =   1380
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
      TabIndex        =   47
      Top             =   4620
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
      TabIndex        =   53
      Top             =   5355
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
      TabIndex        =   49
      Top             =   5250
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
      TabIndex        =   57
      Top             =   6405
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
      TabIndex        =   55
      Top             =   5985
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
      Left            =   5460
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
      Left            =   3990
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
      TabIndex        =   19
      Top             =   6720
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
      TabIndex        =   18
      Top             =   5985
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
      TabIndex        =   26
      Top             =   1785
      Width           =   1905
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
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
      TabIndex        =   25
      Top             =   1785
      Width           =   2010
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   5040
      TabIndex        =   14
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label1 
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
      Height          =   435
      Left            =   105
      TabIndex        =   6
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nota de Venta No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   6510
      TabIndex        =   16
      Top             =   1155
      Width           =   1905
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
      Left            =   6825
      TabIndex        =   4
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label Label3 
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
      Height          =   435
      Left            =   2520
      TabIndex        =   8
      Top             =   630
      Width           =   1065
   End
End
Attribute VB_Name = "Abonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EsNotaDeVenta As Boolean

Private Sub Command1_Click()
  TextoValido TextBanco
  TextoValido TextCheqNo
  TA.Recibi_de = NombreCliente
  Mensajes = "Esta Seguro que desea grabar estos pagos."
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     Calculo_Saldo
     FechaTexto = FechaSistema
     If CheqRecibo.value = 1 Then
        DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
     Else
        DiarioCaja = Val(TxtRecibo)
     End If
     Cta = SinEspaciosIzq(DCBanco.Text)
     Cta1 = SinEspaciosIzq(DCTarjeta.Text)
     NombreBanco1 = Trim(Mid(DCTarjeta.Text, Len(Cta1) + 1, 60))
     NombreBanco1 = Mid(NombreBanco1, 1, 25)
     TA.Cta_CxP = Cta_Cobrar
     TA.CodigoC = CodigoCliente
    'Abono de Factura Caja MN
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Cta_CajaG
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "EFECTIVO MN"
     TA.Cheque = Grupo_No
     TA.Factura = Factura_No
     TA.Abono = TotalCajaMN
     Grabar_Abonos TA
    'Abono de Factura Caja ME
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Cta_CajaGE
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "EFECTIVO MN"
     TA.Cheque = Grupo_No
     TA.Factura = Factura_No
     TA.Abono = TotalCajaME
     Grabar_Abonos TA
    'Abono de Factura Banco
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Trim(SinEspaciosIzq(DCBanco))
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = TextBanco
     TA.Cheque = TextCheqNo
     TA.Factura = Factura_No
     TA.Abono = Total_Bancos
     Grabar_Abonos TA
    'Abono de Factura Tarjeta
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Trim(SinEspaciosIzq(DCTarjeta))
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = NombreBanco1
     TA.Cheque = TextBaucher
     TA.Factura = Factura_No
     TA.Abono = Total_Tarjeta
     Grabar_Abonos TA
    'Abono de Factura Interes Tarjeta
     TA.T = Normal
     TA.TP = "TJ"
     TA.Fecha = MBFecha
     TA.Cta = Cta1
     TA.Cta_CxP = Cta_Tarjetas
     TA.Banco = "INTERES POR TARJETA"
     TA.Cheque = TextBaucher
     TA.Factura = Factura_No
     TA.Abono = Val(TextInteres)
     Grabar_Abonos TA
     Codigo1 = Format(Val(Trim(Mid(TxtSerieRet, 1, 3))), "000")
     Codigo2 = Format(Val(Trim(Mid(TxtSerieRet, 4, 3))), "000")
    'Abono de Factura Rete. IVA Bienes
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Trim(SinEspaciosIzq(DCRetIBienes))
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "RETENCION IVA BIENES"
     TA.Cheque = TextCompRet
     TA.Factura = Factura_No
     TA.Abono = Total_RetIVAB
     TA.AutorizacionR = TxtAutoRet
     TA.Establecimiento = Codigo1
     TA.Emision = Codigo2
     TA.Porcentaje = Val(CBienes)
     Grabar_Abonos_Retenciones TA
    'Abono de Factura Ret IVA Servicio
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Trim(SinEspaciosIzq(DCRetISer))
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "RETENCION IVA SERVICIO"
     TA.Cheque = TextCompRet
     TA.Factura = Factura_No
     TA.Abono = Total_RetIVAS
     TA.AutorizacionR = TxtAutoRet
     TA.Establecimiento = Codigo1
     TA.Emision = Codigo2
     TA.Porcentaje = Val(CServicio)
     Grabar_Abonos_Retenciones TA
    'Abono de Factura Ret. Fuente
     TA.T = Normal
     TA.TP = TipoFactura
     TA.Fecha = MBFecha
     TA.Cta = Trim(SinEspaciosIzq(DCRetFuente))
     TA.Cta_CxP = Cta_Cobrar
     TA.Banco = "RETENCION FUENTE - " & DCCodRet
     TA.Cheque = TextCompRet
     TA.Factura = Factura_No
     TA.Abono = Total_Ret
     TA.AutorizacionR = TxtAutoRet
     TA.Establecimiento = Codigo1
     TA.Emision = Codigo2
     TA.Porcentaje = Val(TextPorc)
     Grabar_Abonos_Retenciones TA
     T = "P"
     If SaldoDisp <= 0 Then
        T = "C"
        SaldoDisp = 0
     End If
     sSQL = "UPDATE Facturas " _
          & "SET Saldo_MN = " & SaldoDisp & ",T = '" & T & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TA.TP & "' " _
          & "AND Serie = '" & TA.Serie & "' " _
          & "AND Autorizacion = '" & TA.Autorizacion & "' " _
          & "AND Factura = " & TA.Factura & " " _
          & "AND CodigoC = '" & TA.CodigoC & "' "
     ConectarAdoExecute sSQL
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
     Imprimir_Comprobante_Caja AdoRecibo, TA
  End If
  If Evaluar Then Unload Abonos
End Sub

Private Sub Command3_Click()
   Unload Abonos
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
   If KeyCode = vbKeyEscape Then Command2.SetFocus
End Sub

Private Sub DCCliente_LostFocus()
  CodigoCliente = Ninguno
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCCliente.Text & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          TA.CI_RUC_Cli = .Fields("CI_RUC")
          NombreCliente = .Fields("Cliente")
          DireccionCli = .Fields("Direccion")
          Grupo_No = .Fields("Grupo")
          'If Evaluar Then DCFactura.Text = Factura_No
       Else
          MsgBox "Cliente no Asignado"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
 'Listamos las autorizaciones de facturas pendientes por cliente
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If Evaluar Then
     sSQL = sSQL _
          & "AND Factura = " & FA.Factura & " " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Autorizacion = '" & FA.Autorizacion & "' "
  Else
     sSQL = sSQL & "AND CodigoC = '" & CodigoCliente & "' "
  End If
  sSQL = sSQL & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "AND NOT TC IN ('OP','C','P') " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDBCombo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
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
       If Not .EOF Then TextPorc = .Fields("Porcentaje")
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
   If KeyCode = vbKeyEscape Then DCCliente.SetFocus
End Sub

Private Sub DCFactura_LostFocus()
  Factura_No = 0: Saldo = 0: Cotizacion = 0: TotalDolar = 0
  Saldo_ME = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura Like '" & DCFactura.Text & "' ")
       If Not .EOF Then
          Factura_No = .Fields("Factura")
          LblObs.Caption = " " & .Fields("Observacion")
          LblNota.Caption = " " & .Fields("Nota")
          Cta_Cobrar = .Fields("Cta_CxP")
          CodigoCliente = .Fields("CodigoC")
          Saldo = Redondear(.Fields("Saldo_MN"), 2)
          Saldo_ME = Redondear(.Fields("Saldo_ME"), 2)
          TipoFactura = .Fields("TC")
          TotalDolar = .Fields("Total_ME")
          Cotizacion = .Fields("Cotizacion")
          TA.Serie = .Fields("Serie")
          TA.Autorizacion = .Fields("Autorizacion")
          If TotalDolar <> 0 Then
             TextRet.Enabled = True
          Else
            TextCajaMN.Enabled = True
            TextCheque.Enabled = True
          End If
          Efectivo = 0: Cheque = 0: Retencion = 0
          LabelDolares.Caption = Format(Cotizacion, "#,##0.00")
          LabelSaldo.Caption = Format(Saldo, "#,##0.00")
          If TotalDolar <> 0 Then
             TextRet.SetFocus
          Else
             TextCajaMN.SetFocus
          End If
          Abonos.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
       End If
    End If
  End With
  Calculo_Saldo
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
End Sub

Private Sub DCRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCRecibo_LostFocus()
  Listar_Clientes_F
End Sub

Private Sub DCTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  TipoFactura = DCTipo
  If TipoFactura = "" Then TipoFactura = Ninguno
  Listar_Facturas_P
  If AdoRecibo.Recordset.RecordCount <= 0 Then CheqRecibo.SetFocus
End Sub

Private Sub Form_Activate()
  Evaluar = True
  With FA
   If .Autorizacion = Ninguno And _
      .Serie = Ninguno And _
      .TC = Ninguno And _
      .Factura = 0 Then Evaluar = False
  End With
  ControlEsNumerico TextCajaMN
  ControlEsNumerico TextCajaME
  ControlEsNumerico TextCheque
  ControlEsNumerico TextTotalBaucher
  If FechaTexto = "" Or FechaTexto = Ninguno Or IsNumeric(FechaTexto) Then FechaTexto = FechaSistema
  MBFecha = FechaTexto
  CBienes.Clear
  CBienes.AddItem "0"
  CBienes.AddItem "30"
  CBienes.AddItem "100"
  CBienes.Text = "0"
  CServicio.Clear
  CServicio.AddItem "0"
  CServicio.AddItem "70"
  CServicio.AddItem "100"
  CServicio.Text = "0"
 'Procesar_Saldo_De_Facturas Abonos, AdoFactura
  
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format(DiarioCaja, "0000000") Else TxtRecibo = ""
  'If Factura_No = 0 Then Command2.Caption = "&Salir"
  Mifecha = BuscarFecha(FechaSistema)
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('BA','CJ') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC,Codigo "
  SelectDBCombo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'TJ' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCTarjeta, AdoTarjeta, sSQL, "NomCuenta"
  
 'Carga los Conceptos de Retención en la Fuente
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CF' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
 'Carga los Conceptos de retención IVA Servicios al DataCombo
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CI' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCRetISer, AdoRetIvaServicio, sSQL, "Cuentas"
 'Carga los Conceptos de retención IVA Bienes al DataCombo
  sSQL = "SELECT (Codigo & '  ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CI' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCRetIBienes, AdoRetIvaBienes, sSQL, "Cuentas"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT * " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFecha) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBFecha) & "# " _
       & "ORDER BY Codigo "
  SelectDBCombo DCCodRet, AdoCodRet, sSQL, "Codigo"
  Grupos_Pendientes
  Listar_Clientes_F
  RatonNormal
  If Evaluar Then
     DCCliente.Text = NombreCliente
     DCCliente.SetFocus
  Else
    'TxtRecibo.SetFocus
  End If
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
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT * " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & BuscarFecha(MBFecha) & "# " _
       & "AND Fecha_Final >= #" & BuscarFecha(MBFecha) & "# " _
       & "ORDER BY Codigo "
  'MsgBox sSQL
  SelectDBCombo DCCodRet, AdoCodRet, sSQL, "Codigo"
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
  TextCajaME.Text = Format(TotalCajaME, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
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
  'TextCajaMN.Text = Format(TextCajaMN.Text, "#,##0.00")
  TextoValido TextCajaMN, True
 'MsgBox TextCajaMN.Text
  TotalCajaMN = Redondear(Val(CCur(TextCajaMN.Text)), 2)
  'TextCajaMN.Text = Format(TotalCajaMN, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
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
  TextCheque.Text = Format(Total_Bancos, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextInteres_GotFocus()
  MarcarTexto TextInteres
End Sub

Private Sub TextInteres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInteres_LostFocus()
  If Mid(TextInteres, Len(TextInteres), 1) = "%" Then
     Valor = Mid(TextInteres, 1, Len(TextInteres) - 1)
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
  TextRecibido = Format(TotalAbonos + Val(CCur(TextInteres)), "#,##0.00")
  MarcarTexto TextRecibido
End Sub

Private Sub TextRecibido_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRecibido_LostFocus()
  TextoValido TextRecibido, True
  Calculo_Saldo
  LabelCambio.Caption = Format(Val(CCur(TextRecibido)) - TotalAbonos - Val(CCur(TextInteres)), "#,##0.00")
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
  TextRet = Format(Total_Ret, "#,##0.00")
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
  TextTotalBaucher.Text = Format(Total_Tarjeta, "#,##0.00")
  Calculo_Saldo
  LabelPend.Caption = Format(SaldoDisp, "#,##0.00")
End Sub

Public Sub Calculo_Saldo()
   TotalAbonos = TotalCajaMN + TotalCajaME + Total_Bancos + Total_Tarjeta + Total_IVA + Total_Ret + Total_RetIVAB + Total_RetIVAS
   SaldoDisp = Saldo - TotalAbonos
End Sub

Public Sub Grupos_Pendientes()
  sSQL = "SELECT C.Grupo " _
       & "FROM Clientes As C,Facturas As F " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T = 'P' " _
       & "AND F.TC IN ('PV','FA','NV') " _
       & "AND F.Saldo_MN > 0 "
  If Evaluar Then
     sSQL = sSQL _
          & "AND F.Factura = " & FA.Factura & " " _
          & "AND F.TC = '" & FA.TC & "' " _
          & "AND F.Serie = '" & FA.Serie & "' " _
          & "AND F.Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND F.CodigoC = '" & FA.CodigoC & "' "
  End If
  sSQL = sSQL & "AND F.CodigoC = C.Codigo " _
       & "GROUP BY C.Grupo " _
       & "ORDER BY C.Grupo "
  SelectDBCombo DCRecibo, AdoRecibo, sSQL, "Grupo"
End Sub

Public Sub Listar_Facturas_P()
  Select Case TipoFactura
    Case "PV": Label2.Caption = " Punto de Venta No."
    Case "NV": Label2.Caption = " Nota de Venta No."
    Case Else: Label2.Caption = " Factura No."
  End Select
  SQL1 = "SELECT F.TC,F.Factura,F.Autorizacion,F.Serie,F.CodigoC,F.Fecha,F.Fecha_V," _
       & "F.Total_MN,F.Total_ME,F.Saldo_MN,F.Saldo_ME,F.Cta_CxP,F.Nota,F.Observacion,F.Cotizacion," _
       & "C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' "
  If Evaluar Then
     SQL1 = SQL1 _
          & "AND F.Factura = " & FA.Factura & " " _
          & "AND F.TC = '" & FA.TC & "' " _
          & "AND F.Serie = '" & FA.Serie & "' " _
          & "AND F.Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND F.CodigoC = '" & FA.CodigoC & "' "
  Else
     SQL1 = SQL1 _
          & "AND F.Autorizacion = '" & Autorizacion & "' " _
          & "AND F.Serie = '" & SerieFactura & "' " _
          & "AND F.TC = '" & TipoFactura & "' " _
          & "AND F.CodigoC = '" & CodigoCliente & "' "
  End If
  SQL1 = SQL1 _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.TC,F.Factura "
  SelectDBCombo DCFactura, AdoFactura, SQL1, "Factura"
  If AdoFactura.Recordset.RecordCount <= 0 Then TxtRecibo.SetFocus
End Sub

Public Sub Listar_Clientes_F()
  Codigo1 = DCRecibo
  If Codigo1 = "" Then Codigo1 = Ninguno
  sSQL = "SELECT Grupo,Cliente,Codigo,CI_RUC,Direccion,Telefono,COUNT(F.Factura) As CantFacturas " _
       & "FROM Clientes As C,Facturas As F " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T = 'P' " _
       & "AND F.TC IN ('PV','FA','NV') " _
       & "AND F.Saldo_MN > 0 " _
       & "AND C.Codigo = F.CodigoC "
  If Evaluar Then
     sSQL = sSQL _
          & "AND F.Factura = " & FA.Factura & " " _
          & "AND F.TC = '" & FA.TC & "' " _
          & "AND F.Serie = '" & FA.Serie & "' " _
          & "AND F.Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND F.CodigoC = '" & FA.CodigoC & "' "
  Else
     If Codigo1 <> Ninguno Then sSQL = sSQL & "AND C.Grupo = '" & Codigo1 & "' "
  End If
  sSQL = sSQL & "GROUP BY Grupo,Cliente,Codigo,CI_RUC,Direccion,Telefono " _
       & "ORDER BY Cliente "
  SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
  If AdoCliente.Recordset.RecordCount <= 0 Then CheqRecibo.SetFocus
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format(Val(TxtRecibo), "0000000")
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
  TextRetIVAB = Format(Total_RetIVAB, "#,##0.00")
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
  TextRetIVAS = Format(Total_RetIVAS, "#,##0.00")
  Calculo_Saldo
End Sub

Private Sub TxtSerieRet_GotFocus()
  MarcarTexto TxtSerieRet
End Sub

Private Sub TxtSerieRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtSerieRet_LostFocus()
  TxtSerieRet = Format(Val(TxtSerieRet), "0000000")
End Sub

Private Sub TxtAutoRet_GotFocus()
  MarcarTexto TxtAutoRet
End Sub

Private Sub TxtAutoRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtAutoRet_LostFocus()
  TextoValido TxtAutoRet
  TxtAutoRet = Format(Val(TxtAutoRet), "0000000000")
End Sub

Private Sub TextCompRet_GotFocus()
  MarcarTexto TextCompRet
End Sub

Private Sub TextCompRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCompRet_LostFocus()
  TextCompRet = Format(Val(TextCompRet), "0000000")
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAutorizacion_LostFocus()
 'Listamos las autorizaciones de facturas pendientes por cliente
  Autorizacion = DCAutorizacion
  If Autorizacion = "" Then Autorizacion = Ninguno
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TC IN ('OP','C','P') "
  If Evaluar Then
     sSQL = sSQL _
          & "AND Factura = " & FA.Factura & " " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Autorizacion = '" & FA.Autorizacion & "' "
  Else
     sSQL = sSQL _
          & "AND CodigoC = '" & CodigoCliente & "' " _
          & "AND Autorizacion = '" & Autorizacion & "' "
  End If
  sSQL = sSQL _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDBCombo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  SerieFactura = DCSerie
  If SerieFactura = "" Then SerieFactura = Ninguno
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TC IN ('OP','C','P') " _
       & "AND CodigoC = '" & CodigoCliente & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Saldo_MN > 0 " _
       & "AND T <> 'A' " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDBCombo DCTipo, AdoDetAcomp, sSQL, "TC"
End Sub

