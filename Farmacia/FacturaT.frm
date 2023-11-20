VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FacturasTours 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   6840
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGPend 
      Height          =   1815
      Left            =   7560
      TabIndex        =   52
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
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
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FacturaT.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   51
      Top             =   4080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3201
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FacturaT.frx":001B
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   120
      TabIndex        =   50
      Top             =   1560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DCGrupo 
      Bindings        =   "FacturaT.frx":0034
      DataSource      =   "AdoCatalogo"
      Height          =   315
      Left            =   4920
      TabIndex        =   49
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   120
      Top             =   4320
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
   Begin VB.TextBox TextVTotal 
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
      Left            =   9660
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "FacturaT.frx":004E
      Top             =   3570
      Width           =   1590
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   8295
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "FacturaT.frx":0053
      Top             =   3570
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
      Left            =   7350
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "FacturaT.frx":0058
      Top             =   3570
      Width           =   960
   End
   Begin VB.TextBox TextRuta 
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
      Left            =   5565
      MaxLength       =   11
      TabIndex        =   33
      Top             =   3570
      Width           =   1800
   End
   Begin VB.TextBox TextIVA 
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
      Left            =   4515
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   40
      Text            =   "FacturaT.frx":005A
      Top             =   6300
      Width           =   1380
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9030
      Picture         =   "FacturaT.frx":005F
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5985
      Width           =   1065
   End
   Begin VB.TextBox TextFechaT 
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
      Left            =   4095
      MaxLength       =   25
      TabIndex        =   23
      Top             =   2835
      Width           =   3270
   End
   Begin VB.TextBox TextCodigo 
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
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   22
      Top             =   2835
      Width           =   1380
   End
   Begin VB.TextBox TextDefinitivo 
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
      MaxLength       =   20
      TabIndex        =   21
      Top             =   2835
      Width           =   2640
   End
   Begin VB.TextBox TextPedidos 
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
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2205
      Width           =   7260
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   5670
      TabIndex        =   15
      Top             =   1575
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#########-#-###"
      Mask            =   "#########-#-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxTelefono 
      Height          =   330
      Left            =   4410
      TabIndex        =   14
      Top             =   1575
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
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin VB.CheckBox CheckGrupo 
      Caption         =   "Factura de Grupo:"
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
      Left            =   2940
      TabIndex        =   4
      Top             =   420
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   2745
      Begin VB.OptionButton OpcMN 
         Caption         =   "Nacional"
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
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton OpcME 
         Caption         =   "Extranjera"
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
         Left            =   1365
         TabIndex        =   2
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.TextBox TextDolar 
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
      Left            =   6300
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "FacturaT.frx":06C9
      Top             =   840
      Width           =   1380
   End
   Begin VB.TextBox TextTicket 
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
      Left            =   4095
      MaxLength       =   12
      TabIndex        =   32
      Top             =   3570
      Width           =   1485
   End
   Begin VB.TextBox TextDescuento 
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
      Left            =   3150
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "FacturaT.frx":06CB
      Top             =   6300
      Width           =   1380
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1785
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   38
      Text            =   "FacturaT.frx":06D0
      Top             =   6300
      Width           =   1380
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
      Height          =   750
      Left            =   10185
      Picture         =   "FacturaT.frx":06D7
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5985
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7875
      Picture         =   "FacturaT.frx":10CD
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5985
      Width           =   1065
   End
   Begin VB.TextBox TextProducto 
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
      Height          =   540
      Left            =   105
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   3570
      Width           =   3900
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   2940
      TabIndex        =   6
      Top             =   840
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   6360
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc AdoChequesPosf 
      Height          =   330
      Left            =   6360
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "ChequesPosf"
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
   Begin MSAdodcLib.Adodc AdoDiarioCaja 
      Height          =   330
      Left            =   4320
      Top             =   5040
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
      Caption         =   "DiarioCaja"
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
      Left            =   4320
      Top             =   4800
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
   Begin MSAdodcLib.Adodc AdoSubCtas1 
      Height          =   330
      Left            =   4320
      Top             =   4560
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
      Caption         =   "SubCtas1"
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
   Begin MSAdodcLib.Adodc AdoSQLs 
      Height          =   330
      Left            =   4320
      Top             =   4320
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
      Caption         =   "SQLs"
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
      Left            =   2160
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc AdoIngresoCaja 
      Height          =   330
      Left            =   120
      Top             =   5040
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
      Caption         =   "IngresoCaja"
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
   Begin MSAdodcLib.Adodc AdoCatalogo 
      Height          =   330
      Left            =   2160
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Catalogo"
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
   Begin MSAdodcLib.Adodc AdoPend 
      Height          =   330
      Left            =   2160
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Pend"
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
   Begin MSAdodcLib.Adodc AdoProductos 
      Height          =   330
      Left            =   2160
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Productos"
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
   Begin MSAdodcLib.Adodc AdoSubCtas 
      Height          =   330
      Left            =   120
      Top             =   4800
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
      Caption         =   "SubCtas"
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
   Begin MSAdodcLib.Adodc AdoAcompañantes 
      Height          =   330
      Left            =   120
      Top             =   4560
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
      Caption         =   "Acompañantes"
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
      Left            =   9660
      TabIndex        =   30
      Top             =   3255
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo Unitario"
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
      TabIndex        =   29
      Top             =   3255
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. PAX"
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
      Left            =   7350
      TabIndex        =   28
      Top             =   3255
      Width           =   960
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5880
      TabIndex        =   44
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Tours"
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
      Left            =   4095
      TabIndex        =   20
      Top             =   2625
      Width           =   3270
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Codigo"
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
      Left            =   2730
      TabIndex        =   19
      Top             =   2625
      Width           =   1380
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura Definitiva:"
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
      Left            =   105
      TabIndex        =   18
      Top             =   2625
      Width           =   2640
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Direccion:"
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
      Left            =   105
      TabIndex        =   16
      Top             =   1995
      Width           =   7260
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " R.U.C./C.I."
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
      Left            =   5670
      TabIndex        =   13
      Top             =   1365
      Width           =   1695
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Teléfono"
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
      Left            =   4410
      TabIndex        =   12
      Top             =   1365
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cliente:"
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
      Left            =   105
      TabIndex        =   11
      Top             =   1365
      Width           =   4320
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COTIZACION:"
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
      Left            =   4935
      TabIndex        =   7
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUTA"
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
      Left            =   5565
      TabIndex        =   27
      Top             =   3255
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TICKET"
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
      Left            =   4095
      TabIndex        =   26
      Top             =   3255
      Width           =   1485
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Facturado"
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
      Left            =   5880
      TabIndex        =   43
      Top             =   5985
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
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
      Left            =   4515
      TabIndex        =   45
      Top             =   5985
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descuento"
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
      TabIndex        =   37
      Top             =   5985
      Width           =   1380
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comision"
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
      TabIndex        =   41
      Top             =   5985
      Width           =   1380
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   2940
      TabIndex        =   3
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   105
      TabIndex        =   46
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
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
      TabIndex        =   42
      Top             =   5985
      Width           =   1695
   End
   Begin VB.Label LabelStockArt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D E T A L L E"
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
      TabIndex        =   25
      Top             =   3255
      Width           =   3900
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Pedido (dd/mm/aaaa):"
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
      TabIndex        =   5
      Top             =   840
      Width           =   2745
   End
   Begin VB.Label LabelFacturaNo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
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
      Height          =   330
      Left            =   9765
      TabIndex        =   10
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura No."
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
      Height          =   330
      Left            =   8190
      TabIndex        =   9
      Top             =   840
      Width           =   1590
   End
End
Attribute VB_Name = "FacturasTours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckGrupo_Click()
  If CheckGrupo.Value = 0 Then
     DCGrupo.Visible = False
  Else
     DCGrupo.Visible = True
  End If
End Sub

Private Sub Command1_Click()
  Unload FacturasTours
End Sub

Private Sub Command2_Click()
   Mensajes = "Esta Seguro que desea grabar: " & Chr(13) & " La Factura No. " & LabelFacturaNo
   Titulo = "Formulario de Grabacion"
   If BoxMensaje = 6 Then
     If OpcME.Value Then Moneda_E = True Else Moneda_E = False
     FechaTexto1 = MBoxFecha.Text
     CalculosTotalesFactura
     Total_Factura = Round(Total_Con_IVA - Total_Desc - Total_Comision + Total_IVA)
     LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
     Total_FacturaME = 0
     If Moneda_E Then
        Total_FacturaME = Total_Factura
        Total_Factura = Round(Total_Factura * Dolar)
     Else
        Total_Factura = Round(Total_Factura)
     End If
     ProcGrabar
   End If
End Sub

Private Sub Command3_Click()
Dim NivelNo1
Dim TipoProc1
Dim TipoFacturas1
  Titulo = "Imprimir Factura"
  Mensaje = "Ingrese el Numero de Factura:"
  Factura_No = Val(InputBox(Mensaje, Titulo))
  If Factura_No > 0 Then
     NivelNo1 = NivelNo
     TipoProc1 = TipoProc
     TipoFacturas1 = TipoFacturas
     SQL1 = "SELECT Facturas.*,Clientes.* "
     SQL1 = SQL1 & "FROM Facturas,Clientes "
     SQL1 = SQL1 & "WHERE Factura = " & Factura_No & " "
     SQL1 = SQL1 & "AND Clientes.Codigo = Codigo_C "
     SelectData AdoFactura, SQL1, False
     If AdoFactura.Recordset.RecordCount > 0 Then
        NivelNo = AdoFactura.Recordset.Fields("Nivel")
        TipoFacturas = AdoFactura.Recordset.Fields("Nota")
        SQL2 = "SELECT * FROM Niveles "
        SQL2 = SQL2 & "WHERE Nivel = '" & NivelNo & "' "
        SelectData AdoSubCtas1, SQL2, False
        If AdoSubCtas1.Recordset.RecordCount > 0 Then
           TipoProc = AdoSubCtas1.Recordset.Fields("TP")
        End If
     End If
    'Consultamos el detalle de la Factura
     SQL2 = "SELECT * FROM Detalle_Factura "
     SQL2 = SQL2 & "WHERE Factura_No = " & Factura_No & " "
     SelectData AdoDetalle, SQL2, False
     SQL3 = "SELECT * FROM Acompañantes "
     SQL3 = SQL3 & "WHERE Factura = " & Factura_No & " "
     SelectData AdoDetAcomp, SQL3, False
     ImprimirFactTurs AdoFactura, AdoDetAcomp, AdoDetalle
     NivelNo = NivelNo1
     TipoProc = TipoProc1
     TipoFacturas = TipoFacturas1
     Select Case TipoProc
       Case "FA":
            sSQL = "SELECT PRODUCTO,CANT,PRECIO AS V_UNIT,TOTAL "
       Case "FA1":
            sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS TARIFA_IVA_TASA "
       Case "FA2":
            sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS COMISION "
       Case "FA3":
            sSQL = "SELECT PRODUCTO AS NOMBRES,PRECIO AS VALOR_UNIT,TOTAL AS VALOR_TOTAL "
     End Select
     sSQL = sSQL & "FROM Detalles_" & CodigoUsuario & " "
     SelectDataGrid DGDetalle, AdoProductos, sSQL
     OpcMN.SetFocus
  End If
End Sub

Private Sub DCCliente_GotFocus()
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE E='C' "
   sSQL = sSQL & "ORDER BY Grupo,Cliente "
   SelectData AdoCliente, sSQL
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
   DCCliente.Text = UCase(DCCliente.Text)
   Empleados = False
   CodigoCli = Ninguno
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE Cliente = '" & DCCliente.Text & "' "
   AdoCliente.RecordSource = sSQL: AdoCliente.Refresh
   If AdoCliente.Recordset.RecordCount > 0 Then
      SaldoPendiente = 0
      DCCliente.Text = AdoCliente.Recordset.Fields("Cliente")
      MBoxTelefono.Text = AdoCliente.Recordset.Fields("Telefono")
      CodigoCli = AdoCliente.Recordset.Fields("Codigo")
      DireccionCli = AdoCliente.Recordset.Fields("Direccion")
      MBoxRUC.Text = AdoCliente.Recordset.Fields("RUC_CI")
      TextPedidos.Text = DireccionCli
      'Empleados = AdoCliente.Recordset.Fields("E")
   Else
      Mensajes = "Este Cliente No existe," & Chr(13)
      Mensajes = Mensajes & "Repita la operacion."
      MsgBox Mensajes
      MBoxTelefono.Text = "00-000-000"
      MBoxRUC.Text = "000000000-0-000"
      TextPedidos.Text = ""
      MBoxTelefono.SetFocus
   End If
End Sub

Private Sub DGDetalle_AfterDelete()
  CalculosTotalesFactura
End Sub

Private Sub DGDetalle_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoProductos)
End Sub

Private Sub Form_Activate()
   CTDetalles_F
   Label9.Caption = Cta_General & " " & TipoFacturas
   Cta = Codigo
   Mifecha = BuscarFecha(FechaSistema)
   NumComp = ReadSetDataNum("Facturas ", True, False)
   Factura_No = NumComp
   sSQL = "DELETE * FROM Det_Acomp "
   ConectarAdoExecute sSQL
   'DeleteData AdoDetalle, sSQL
   sSQL = "SELECT * FROM Det_Acomp "
   SelectData AdoDetAcomp, sSQL, False
   sSQL = "SELECT * FROM Acompañantes "
   SelectData AdoAcompañantes, sSQL, False
   sSQL = "SELECT * FROM Clientes ORDER BY Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente", False
   sSQL = "SELECT Codigo & Space(10) & Cuenta As CtaGrupo "
   sSQL = sSQL & "FROM Catalogo "
   sSQL = sSQL & "WHERE TC = 'G' AND DG = 'D' "
   sSQL = sSQL & "ORDER BY Codigo "
   SelectDBCombo DCGrupo, AdoCatalogo, sSQL, "CtaGrupo", False
   sSQL = "DELETE * FROM Detalles" _
        & "WHERE CodigoU = '" & CodigoUsuario & " "
   ConectarAdoExecute sSQL
   'DeleteData AdoDetalle, sSQL
   Select Case TipoProc
     Case "FA":
          sSQL = "SELECT PRODUCTO,CANT,PRECIO AS V_UNIT,TOTAL "
          Label5.Visible = False
          Label7.Visible = False
          TextTicket.Visible = False
          TextRuta.Visible = False
          TextVTotal.Enabled = False
     Case "FA1":
          sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS TARIFA_IVA_TASA "
          LabelStockArt.Caption = "N O M B R E S"
          Label19.Caption = "TARIFA"
          Label20.Caption = "TARIFA_IVA_TASA"
          TextCant.Visible = False
          Label16.Visible = False
          DGPend.Visible = False
          Label11.Visible = False
          Label14.Visible = False
          Label15.Visible = False
          TextDefinitivo.Visible = False
          TextCodigo.Visible = False
          TextFechaT.Visible = False
     Case "FA2":
          sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS COMISION "
          LabelStockArt.Caption = "N O M B R E S"
          Label19.Caption = "TARIFA"
          Label20.Caption = "COMISION"
          TextCant.Visible = False
          Label16.Visible = False
          DGPend.Visible = False
          Label11.Visible = False
          Label14.Visible = False
          Label15.Visible = False
          TextDefinitivo.Visible = False
          TextCodigo.Visible = False
          TextFechaT.Visible = False
       Case "FA3":
          sSQL = "SELECT PRODUCTO AS NOMBRES,PRECIO AS VALOR_UNIT,TOTAL AS VALOR_TOTAL "
          LabelStockArt.Caption = "N O M B R E S"
          Label19.Caption = "VALOR UNIT."
          Label20.Caption = "VALOR TOTAL"
          TextCant.Visible = False
          TextTicket.Visible = False
          TextRuta.Visible = False
          Label7.Visible = False
          Label5.Visible = False
          Label16.Visible = False
          DGPend.Visible = False
          Label11.Visible = False
          Label14.Visible = False
          Label15.Visible = False
          TextDefinitivo.Visible = False
          TextCodigo.Visible = False
          TextFechaT.Visible = False
   End Select
   sSQL = sSQL & "FROM Detalles_" & CodigoUsuario & " "
   SelectDataGrid DGDetalle, AdoProductos, sSQL
   LabelFacturaNo.Caption = Format(NumComp, "000000")
   sSQL = "SELECT Fecha FROM Diario_Caja "
   sSQL = sSQL & "WHERE T = '" & Normal & "' "
   sSQL = sSQL & "AND Fecha <> #" & Mifecha & "# "
   sSQL = sSQL & "ORDER BY Fecha "
   SelectData AdoDiarioCaja, sSQL, False
   RatonNormal
   With AdoDiarioCaja.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Cadena = "Falta cerrar El Diario de Caja, Fecha: " & .Fields("Fecha") & "," & Chr(13)
       Cadena = Cadena & "hasta mientras podrá seguir facturando," & Chr(13)
       Cadena = Cadena & "pero tenga en cuenta este mensaje."
       MsgBox Cadena
   End If
   End With
 RatonNormal
 OpcMN.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FacturasTours
  'Abriendo bases relacionadas
   ConectarAdodc AdoCliente
   ConectarAdodc AdoProductos
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoAcompañantes
   ConectarAdodc AdoDiarioCaja
   ConectarAdodc AdoChequesPosf
   ConectarAdodc AdoIngresoCaja
   ConectarAdodc AdoFactura
   ConectarAdodc AdoDetalle
   ConectarAdodc AdoPend
   ConectarAdodc AdoSubCtas1
   ConectarAdodc AdoSQLs
   ConectarAdodc AdoSubCtas
   ConectarAdodc AdoCatalogo
   TextCant.Text = ""
   TextVUnit.Text = ""
   TextVTotal.Text = ""
   Modificar = False
   Bandera = True
End Sub

Private Sub MBoxFecha_GotFocus()
  If OpcME.Value Then Moneda_E = True Else Moneda_E = False
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, True
   FechaTexto1 = MBoxFecha.Text
End Sub

Private Sub TextCant_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   TextVTotal.Text = Format(Real1, "#,##0.00")
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
   If TextCant.Text = "" Then TextCant.Text = "0"
   Cantidad = Val(TextCant.Text)
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   TextVTotal.Text = Format(Real1, "#,##0.00")
End Sub

Private Sub TextCodigo_GotFocus()
  TextCodigo.Text = ""
End Sub

Private Sub TextCodigo_LostFocus()
  If TextCodigo.Text = "" Then TextCodigo.Text = Ninguno
End Sub

Private Sub TextComision_GotFocus()
  CalculosTotalesFactura
  TextComision.Text = ""
End Sub

Private Sub TextComision_LostFocus()
  If TextComision.Text = "" Then TextComision.Text = "0.00"
  Total_Comision = Round(Val(TextComision.Text) * Total_Sin_IVA / 100)
  TextComision.Text = Format(Total_Comision, "#,##0.00")
End Sub

Private Sub TextDefinitivo_GotFocus()
  TextDefinitivo.Text = ""
End Sub

Private Sub TextDefinitivo_LostFocus()
  If TextDefinitivo.Text = "" Then TextDefinitivo.Text = Ninguno
End Sub

Private Sub TextDescuento_GotFocus()
  TextDescuento.Text = ""
End Sub

Private Sub TextDescuento_LostFocus()
  If TextDescuento.Text = "" Then TextDescuento.Text = "0.00"
  If TipoProc = "FA1" Then
     Total_Desc = Round(Val(TextDescuento.Text) * Total_Sin_IVA / 100)
  Else
     Total_Desc = Round(Val(TextDescuento.Text) * Total_Con_IVA / 100)
  End If
  TextDescuento.Text = Format(Total_Desc, "#,##0.00")
End Sub

Private Sub TextDolar_GotFocus()
  TextDolar.Text = Dolar
End Sub

Private Sub TextDolar_LostFocus()
  If Val(TextDolar.Text) <= 0 Then TextDolar.Text = Dolar
  Dolar = Val(TextDolar.Text)
End Sub

Private Sub TextFechaT_GotFocus()
  TextFechaT.Text = ""
End Sub

Private Sub TextFechaT_LostFocus()
  If TextFechaT.Text = "" Then TextFechaT.Text = Ninguno
End Sub

Private Sub TextIVA_GotFocus()
  TextIVA.Text = ""
End Sub

Private Sub TextIVA_LostFocus()
  CalculosTotalesFactura
  If TextIVA.Text = "" Then TextIVA.Text = "0.00"
  If TipoProc = "FA1" Then
     If Total_Comision <> 0 Then
        Total_IVA = Round(Val(TextIVA.Text) * Total_Comision / 100)
     Else
        Total_IVA = Round(Val(TextIVA.Text) * (Total_Sin_IVA - Total_Desc) / 100)
     End If
  Else
  If Total_Comision <> 0 Then
     Total_IVA = Round(Val(TextIVA.Text) * Total_Comision / 100)
  Else
     Total_IVA = Round(Val(TextIVA.Text) * (Total_Con_IVA - Total_Desc) / 100)
  End If
  End If
  'MsgBox Total_Desc
  TextIVA.Text = Format(Total_IVA, "#,##0.00")
  Total_Factura = Round(Total_Con_IVA - Total_Desc - Total_Comision + Total_IVA)
  LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
End Sub

Private Sub TextPedidos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPedidos_LostFocus()
   TextoValido TextPedidos
End Sub

Private Sub TextProducto_GotFocus()
   TextProducto.Text = ""
End Sub

Private Sub TextProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then TextComision.SetFocus
End Sub

Private Sub TextProducto_LostFocus()
   If TextProducto.Text = "" Then TextProducto.Text = Ninguno
End Sub

Private Sub TextRuta_LostFocus()
  If TextRuta.Text = "" Then TextRuta.Text = Ninguno
  TextRuta.Text = UCase(TextRuta.Text)
  TextVTotal.Text = ""
End Sub

Private Sub TextTicket_GotFocus()
  If Len(TextTicket.Text) > 2 Then
     TextTicket.Text = Mid(TextTicket.Text, 1, 2) & Val(Mid(TextTicket.Text, 3, Len(TextTicket.Text)) + 1)
  End If
End Sub

Private Sub TextTicket_LostFocus()
  If TextTicket.Text = "" Then TextTicket.Text = Ninguno
  TextTicket.Text = UCase(TextTicket.Text)
End Sub

Private Sub TextVTotal_LostFocus()
   TextoValido TextVTotal, True
   If AdoProductos.Recordset.RecordCount <= 15 Then
      If TextProducto.Text <> Ninguno Then
         Total = Val(TextVTotal.Text)
         AdoProductos.Recordset.AddNew
         Select Case TipoProc
           Case "FA1"
                AdoProductos.Recordset.Fields("NOMBRES") = TextProducto.Text
                AdoProductos.Recordset.Fields("TICKET") = TextTicket.Text
                AdoProductos.Recordset.Fields("RUTA") = TextRuta.Text
                AdoProductos.Recordset.Fields("TARIFA") = ValorUnit
                AdoProductos.Recordset.Fields("TARIFA_IVA_TASA") = Total
           Case "FA2"
                AdoProductos.Recordset.Fields("NOMBRES") = TextProducto.Text
                AdoProductos.Recordset.Fields("RUTA") = TextRuta.Text
                AdoProductos.Recordset.Fields("TARIFA") = ValorUnit
                AdoProductos.Recordset.Fields("COMISION") = Total
           Case "FA3"
                AdoProductos.Recordset.Fields("NOMBRES") = TextProducto.Text
                AdoProductos.Recordset.Fields("VALOR_UNIT") = ValorUnit
                AdoProductos.Recordset.Fields("VALOR_TOTAL") = Total
         End Select
         AdoProductos.Recordset.Update
         CalculosTotalesFactura
         TextVTotal.Text = ""
         TextVUnit.Text = ""
         'TextVTotal.Text = Format(Total, "#,##0.00")
      End If
   Else
      MsgBox "Ya no se puede ingresar mas datos."
   End If
   TextProducto.SetFocus
End Sub

Private Sub TextVUnit_Change()
  If TipoProc = "FA" Then TextVTotal.Text = Format(Cantidad * Val(TextVUnit.Text), "#,##0.00")
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
   If TextVUnit.Text = "" Then TextVUnit.Text = "0"
   ValorUnit = Val(TextVUnit.Text)
   If TipoProc = "FA" And TextProducto.Text <> Ninguno Then
      Total = Cantidad * ValorUnit
      If AdoProductos.Recordset.RecordCount <= 15 Then
         AdoProductos.Recordset.AddNew
         AdoProductos.Recordset.Fields("PRODUCTO") = TextProducto.Text
         AdoProductos.Recordset.Fields("CANT") = Cantidad
         AdoProductos.Recordset.Fields("V_UNIT") = ValorUnit
         AdoProductos.Recordset.Fields("TOTAL") = Total
         AdoProductos.Recordset.Update
         CalculosTotalesFactura
         TextVUnit.Text = ""
         TextVTotal.Text = Format(Total, "#,##0.00")
      Else
         MsgBox "Ya no se puede ingresar mas datos."
      End If
      TextProducto.SetFocus
   End If
End Sub

Public Sub CalculosTotalesFactura()
   TextoValido TextIVA, True
   TextoValido TextDescuento, True
   TextoValido TextComision, True
   If TipoProc = "FA" Then
      sSQL = "UPDATE Detalles_" & CodigoUsuario & " "
      sSQL = sSQL & "SET TOTAL = CANT * PRECIO "
      'UpdateData AdoProductos, sSQL
   End If
   Total_Factura = 0: Total_Con_IVA = 0: Total_Sin_IVA = 0
   With AdoProductos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not AdoProductos.Recordset.EOF
          Select Case TipoProc
            Case "FA"
               Total_Con_IVA = Total_Con_IVA + .Fields("TOTAL")
               Total_Sin_IVA = Total_Sin_IVA + .Fields("V_UNIT")
            Case "FA1"
               Total_Con_IVA = Total_Con_IVA + .Fields("TARIFA_IVA_TASA")
               Total_Sin_IVA = Total_Sin_IVA + .Fields("TARIFA")
            Case "FA2"
               Total_Con_IVA = Total_Con_IVA + .Fields("COMISION")
               Total_Sin_IVA = Total_Sin_IVA + .Fields("TARIFA")
            Case "FA3"
               Total_Con_IVA = Total_Con_IVA + .Fields("VALOR_TOTAL")
               Total_Sin_IVA = Total_Sin_IVA + .Fields("VALOR_UNIT")
          End Select
         .MoveNext
       Loop
   End If
   End With
   'Total_Factura = Total_Con_IVA
   Total_IVA = CDbl(TextIVA.Text)
   Total_Desc = CDbl(TextDescuento.Text)
   Total_Comision = CDbl(TextComision.Text)
   Total_Factura = Total_Con_IVA - Total_Desc - Total_Comision + Total_IVA
   LabelConIVA.Caption = Format(Total_Con_IVA, "#,##0.00")
   LabelTotal.Caption = Format(Total_Factura, "#,##0.00")
   TextCant.Text = ""
   TextVTotal.Text = ""
End Sub

Public Sub ProcGrabar()
  If CodigoCli = Ninguno Then
     With AdoCliente.Recordset
         .AddNew
          Numero = ReadSetDataNum("Clientes", True, True)
          CodigoCli = FormatoCodigo(DCCliente.Text, Numero)
         .Fields("T") = "C"
         .Fields("Codigo") = CodigoCli
         .Fields("Cliente") = DCCliente.Text
         .Fields("Empresa") = DCCliente.Text
         .Fields("Direccion") = TextPedidos.Text
         .Fields("Telefono") = MBoxTelefono.Text
         .Fields("Celular") = MBoxTelefono.Text
         .Fields("FAX") = MBoxTelefono.Text
         .Fields("RUC_CI") = MBoxRUC.Text
         .Fields("Ciudad") = Ninguno
         .Update
     End With
  End If
  If TextDefinitivo.Text = "" Then TextDefinitivo.Text = Ninguno
  If TextCodigo.Text = "" Then TextCodigo.Text = Ninguno
  If TextFechaT.Text = "" Then TextFechaT.Text = Ninguno
 'Seteamos los encabezados para las facturas
  If AdoProductos.Recordset.RecordCount > 0 Then
     RatonReloj
     NumComp = ReadSetDataNum("Facturas", True, True)
     SelectData AdoFactura, "Facturas", False
     SelectData AdoDetalle, "Detalle_Factura", False
     FechaTexto = MBoxFecha.Text
     Factura_No = NumComp
     CalculosTotalesFactura
     Total_Factura = Round(Total_Con_IVA - Total_Desc - Total_Comision + Total_IVA)
     Total_FacturaME = 0
     Saldo = Total_Factura
     If Moneda_E Then
        Total_FacturaME = Total_Factura
        Total_Factura = Round(Total_Factura * Dolar)
     End If
     Saldo_ME = Total_FacturaME
     SelectData AdoDiarioCaja, "Diario_Caja ", False
     DiarioCaja = 0
     With AdoDiarioCaja.Recordset
         .AddNew
         .Fields("Banco") = Ninguno
         .Fields("Cheque") = Ninguno
         .Fields("Abonos_ME") = 0
         .Fields("Saldo_ME") = 0
         .Fields("T") = Normal
         .Fields("TP") = ventas
         .Fields("Fecha") = FechaTexto
         .Fields("Diario_No") = DiarioCaja
         .Fields("Caja_No") = IngresoCaja
         .Fields("Factura") = Factura_No
         .Fields("Monto_ME") = Total_FacturaME
         .Fields("Monto_MN") = Total_Factura
         .Fields("Caja_ME") = 0
         .Fields("Caja_MN") = 0
         .Fields("Caja_Vaucher") = 0
         .Fields("Abonos_MN") = 0
         .Fields("Saldo_MN") = Saldo
         .Fields("Codigo_C") = CodigoCli
         .Fields("CtaxCob") = Cta_General
         .Fields("CtaxVent") = Cta_Ventas
          If Moneda_E Then
            .Fields("Cotizacion") = Dolar
            .Fields("Abonos_ME") = 0
            .Fields("Saldo_ME") = Saldo / Dolar
          Else
            .Fields("Cotizacion") = 0
          End If
         .Update
     End With
     If Saldo < 0 Then Saldo = 0
     If Saldo > 0 Then
        TextoFormaPago = PagoCred
        T = Pendiente
     Else
        TextoFormaPago = PagoCont
        T = Cancelado
     End If
     Saldo_ME = 0
     If Moneda_E Then Saldo_ME = Saldo / Dolar
     AdoFactura.Recordset.AddNew
     AdoFactura.Recordset.Fields("T") = T
     AdoFactura.Recordset.Fields("ME") = Moneda_E
     AdoFactura.Recordset.Fields("Cta_CxC") = Cta_General
     AdoFactura.Recordset.Fields("Cta_Venta") = Cta_Ventas
     AdoFactura.Recordset.Fields("Nivel") = NivelNo
     AdoFactura.Recordset.Fields("Factura") = Factura_No
     AdoFactura.Recordset.Fields("Fecha") = FechaTexto
     AdoFactura.Recordset.Fields("Fecha_C") = FechaTexto
     AdoFactura.Recordset.Fields("Fecha_V") = FechaTexto
     AdoFactura.Recordset.Fields("Codigo_C") = CodigoCli
     AdoFactura.Recordset.Fields("Forma_Pago") = TextoFormaPago
     AdoFactura.Recordset.Fields("Comision") = Total_Comision
     AdoFactura.Recordset.Fields("Descuento") = Total_Desc
     AdoFactura.Recordset.Fields("IVA") = Total_IVA
     AdoFactura.Recordset.Fields("SubTotal") = Total_Con_IVA
     AdoFactura.Recordset.Fields("Total_ME") = Total_FacturaME
     AdoFactura.Recordset.Fields("Total_MN") = Total_Factura
     AdoFactura.Recordset.Fields("Saldo_MN") = Saldo
     AdoFactura.Recordset.Fields("Saldo_ME") = Saldo_ME
     AdoFactura.Recordset.Fields("Definitivo") = TextDefinitivo.Text
     AdoFactura.Recordset.Fields("Codigo_T") = TextCodigo.Text
     AdoFactura.Recordset.Fields("Fecha_Tours") = TextFechaT.Text
     AdoFactura.Recordset.Fields("Vendedor") = NombreUsuario
     AdoFactura.Recordset.Fields("Cod_Ejec") = Ninguno
     AdoFactura.Recordset.Fields("Contrato_No") = Ninguno
     AdoFactura.Recordset.Fields("Nota") = TipoFacturas
     AdoFactura.Recordset.Fields("Observacion") = Ninguno
     If Moneda_E Then
        AdoFactura.Recordset.Fields("Cotizacion") = Dolar
     Else
        AdoFactura.Recordset.Fields("Cotizacion") = 0
     End If
     AdoFactura.Recordset.Update
     sSQL = "SELECT * FROM Det_Acomp "
     SelectData AdoDetAcomp, sSQL, False
     With AdoDetAcomp.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          Do While Not AdoDetAcomp.Recordset.EOF
             AdoAcompañantes.Recordset.AddNew
             AdoAcompañantes.Recordset.Fields("No_Hab") = Ninguno
             AdoAcompañantes.Recordset.Fields("Factura") = Factura_No
             AdoAcompañantes.Recordset.Fields("Acompañante") = .Fields("Acompañante")
             AdoAcompañantes.Recordset.Update
            .MoveNext
          Loop
      End If
     End With
     With AdoProductos.Recordset
         .MoveFirst
          Do While Not .EOF
             AdoDetalle.Recordset.AddNew
             AdoDetalle.Recordset.Fields("T") = T
             AdoDetalle.Recordset.Fields("Codigo") = Ninguno
             AdoDetalle.Recordset.Fields("CodigoL") = Ninguno
             AdoDetalle.Recordset.Fields("No_Hab") = Ninguno
             AdoDetalle.Recordset.Fields("Cod_Ejec") = Ninguno
             AdoDetalle.Recordset.Fields("Codigo_C") = CodigoCli
             AdoDetalle.Recordset.Fields("Factura_No") = Factura_No
             AdoDetalle.Recordset.Fields("Codigo_C") = CodigoCli
             AdoDetalle.Recordset.Fields("Fecha") = FechaTexto
            'Blancos
             AdoDetalle.Recordset.Fields("Ticket") = Ninguno
             AdoDetalle.Recordset.Fields("Ruta") = Ninguno
             AdoDetalle.Recordset.Fields("Producto") = Ninguno
             AdoDetalle.Recordset.Fields("Cantidad") = 0
             Select Case TipoProc
               Case "FA"
                    AdoDetalle.Recordset.Fields("Producto") = .Fields("PRODUCTO")
                    AdoDetalle.Recordset.Fields("Cantidad") = .Fields("CANT")
                    AdoDetalle.Recordset.Fields("Precio") = .Fields("V_UNIT")
                    AdoDetalle.Recordset.Fields("Total") = .Fields("TOTAL")
               Case "FA1"
                    AdoDetalle.Recordset.Fields("Ticket") = .Fields("TICKET")
                    AdoDetalle.Recordset.Fields("Ruta") = .Fields("RUTA")
                    AdoDetalle.Recordset.Fields("Producto") = .Fields("NOMBRES")
                    AdoDetalle.Recordset.Fields("Precio") = .Fields("TARIFA")
                    AdoDetalle.Recordset.Fields("Total") = .Fields("TARIFA_IVA_TASA")
               Case "FA2"
                    AdoDetalle.Recordset.Fields("Ruta") = .Fields("RUTA")
                    AdoDetalle.Recordset.Fields("Producto") = .Fields("NOMBRES")
                    AdoDetalle.Recordset.Fields("Precio") = .Fields("TARIFA")
                    AdoDetalle.Recordset.Fields("Total") = .Fields("COMISION")
               Case "FA3"
                    AdoDetalle.Recordset.Fields("Producto") = .Fields("NOMBRES")
                    AdoDetalle.Recordset.Fields("Precio") = .Fields("VALOR_UNIT")
                    AdoDetalle.Recordset.Fields("Total") = .Fields("VALOR_TOTAL")
             End Select
             AdoDetalle.Recordset.Update
            .MoveNext
          Loop
     End With
    'Grabamos si es factura de grupo
     If CheckGrupo.Value <> 0 Then
        Codigo = Trim("FA" & Format(Factura_No, "000000"))
        With Co
            .Fecha = FechaTexto
            .Cotizacion = Dolar
            .Usuario = NombreUsuario
            .Concepto = "Credito de Factura No. " & Format(Factura_No, "000000")
            .CodigoB = Codigo
            .RUC_CI = "000000000-0-000"
            .TP = "FA"
            .Numero = Factura_No
            .Monto_Total = Total_Factura
            .Efectivo = Total_Factura
             'If .CEj = "" Then .CEj = Normal
             If .T = "" Then .T = Normal
        End With
        SQL1 = "SELECT * FROM Comprobantes "
        SelectData AdoSubCtas, SQL1, False
        With AdoSubCtas.Recordset
            .AddNew
            .Fields("T") = Co.T
            '.Fields("CEj") = Co.CEj
            .Fields("Fecha") = Co.Fecha
            .Fields("TP") = Co.TP
            .Fields("Numero") = Co.Numero
            .Fields("Beneficiario") = Co.CodigoB
            .Fields("RUC_CI") = Co.RUC_CI
            .Fields("Monto_Total") = Round(Co.Monto_Total)
            .Fields("Concepto") = Co.Concepto
            .Fields("Efectivo") = Co.Efectivo
            .Fields("Cotizacion") = Co.Cotizacion
            .Fields("Usuario") = Co.Usuario
            .Update
        End With
       'Grabamos SubCtas
        SQL1 = "SELECT * FROM TransaccionesSC "
        SelectData AdoSubCtas, SQL1, False
        With AdoSubCtas.Recordset
            .AddNew
            .Fields("T") = Normal
            .Fields("TC") = "G"
            .Fields("Cta") = SinEspaciosIzq(DCGrupo.Text)
            .Fields("Codigo") = Co.CodigoB
            .Fields("Fecha") = Co.Fecha
            .Fields("Fecha_V") = Co.Fecha
            .Fields("TP") = Co.TP
            .Fields("Numero") = Co.Numero
            .Fields("Factura") = Co.Numero
            .Fields("Debitos") = 0
            .Fields("Creditos") = Total_Factura
            .Fields("Debitos_ME") = 0
            .Fields("Creditos_ME") = Total_FacturaME
            .Fields("Saldo") = 0
            .Fields("Saldo_ME") = 0
            .Update
        End With
        SQL1 = "DELETE * FROM Beneficiarios " _
             & "WHERE Codigo = '" & Codigo & "' "
        'DeleteData AdoSubCtas, SQL1
        sSQL = "SELECT * FROM Beneficiarios "
        sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
        sSQL = sSQL & "AND TC = 'F' "
        SelectData AdoSubCtas, sSQL, False
        With AdoSubCtas.Recordset
            .AddNew
            .Fields("TC") = "F"
            .Fields("Codigo") = Codigo
            .Fields("Beneficiario") = "Factura No. " & Format(Factura_No, "000000")
            .Fields("Ciudad") = Ninguno
            .Fields("Direccion") = Ninguno
            .Fields("Presupuesto") = 0
            .Fields("RUC_CI") = "000000000-0-000"
            .Fields("Telefono") = "00-000-000"
            .Fields("Celular") = "00-000-000"
            .Fields("FAX") = "00-000-000"
            .Update
        End With
     End If
     NumComp = ReadSetDataNum("Facturas", True, False)
     'Grabamos el numero de factura
     sSQL = "DELETE * FROM Detalles_" & CodigoUsuario & " "
     ConectarAdoExecute sSQL
     Select Case TipoProc
       Case "FA": sSQL = "SELECT PRODUCTO,CANT,PRECIO AS V_UNIT,TOTAL "
       Case "FA1": sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS TARIFA_IVA_TASA "
       Case "FA2": sSQL = "SELECT PRODUCTO AS NOMBRES, TICKET, RUTA,PRECIO AS TARIFA,TOTAL AS COMISION "
       Case "FA3": sSQL = "SELECT PRODUCTO AS NOMBRES, PRECIO AS VALOR_UNIT,TOTAL AS VALOR_TOTAL "
     End Select
     sSQL = sSQL & "FROM Detalles_" & CodigoUsuario & " "
     SelectDataGrid DGDetalle, AdoProductos, sSQL
     SQL1 = "SELECT Facturas.*,Clientes.* "
     SQL1 = SQL1 & "FROM Facturas,Clientes "
     SQL1 = SQL1 & "WHERE Factura = " & Factura_No & " "
     SQL1 = SQL1 & "AND Clientes.Codigo = Codigo_C "
     'Consultamos el detalle de la Factura
     SQL2 = "SELECT * FROM Detalle_Factura "
     SQL2 = SQL2 & "WHERE Factura_No = " & Factura_No & " "
     SelectData AdoFactura, SQL1, False
     SelectData AdoDetalle, SQL2, False
     sSQL = "DELETE * FROM Det_Acomp "
     'DeleteData AdoDetalle, sSQL
     sSQL = "SELECT * FROM Det_Acomp "
     SelectData AdoDetAcomp, sSQL, False
     sSQL = "SELECT * FROM Clientes "
     sSQL = sSQL & "WHERE E='C' "
     sSQL = sSQL & "ORDER BY Grupo,Cliente "
     SelectData AdoCliente, sSQL
     RatonNormal
     
     Mensajes = "Pago al Contado"
     Titulo = "Formulario de Pago"
     TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
     If J = 6 Then
        Cadena = DCCliente.Text
        FPagoContado.MBoxFecha.Text = MBoxFecha.Text
        FPagoContado.Show 1
     End If
     
     ImprimirFactTurs AdoFactura, AdoDetAcomp, AdoDetalle
     NumComp = ReadSetDataNum("Facturas", True, False)
     LabelFacturaNo = Format(NumComp, "000000")
  Else
     MsgBox "No se puede grabar la Factura," & Chr(13) & "falta datos."
  End If
End Sub

