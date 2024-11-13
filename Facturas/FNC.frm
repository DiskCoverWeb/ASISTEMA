VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FNotasDeCredito 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANULACION DE FACTURAS"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCContraCta 
      Bindings        =   "FNC.frx":0000
      DataSource      =   "AdoContraCta"
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Top             =   840
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cta"
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
   Begin VB.TextBox TextCompRet 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6405
      MaxLength       =   9
      TabIndex        =   11
      Text            =   "00000000"
      Top             =   840
      Width           =   1170
   End
   Begin VB.TextBox TextCheqNo 
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
      Height          =   360
      Left            =   5565
      MaxLength       =   8
      TabIndex        =   9
      Text            =   "001001"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox TxtConIVA 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   5880
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "FNC.frx":001B
      Top             =   6825
      Width           =   1905
   End
   Begin VB.TextBox TxtSinIVA 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   5880
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   55
      Text            =   "FNC.frx":001F
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Frame FrmProductos 
      Height          =   1380
      Left            =   105
      TabIndex        =   28
      Top             =   2415
      Width           =   14925
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   10500
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "FNC.frx":0023
         Top             =   945
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   9555
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "FNC.frx":0028
         Top             =   945
         Width           =   960
      End
      Begin VB.TextBox TextDesc 
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   11865
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   40
         Text            =   "FNC.frx":002A
         Top             =   945
         Width           =   1380
      End
      Begin MSDataListLib.DataCombo DCArticulo 
         Bindings        =   "FNC.frx":002F
         DataSource      =   "AdoArticulo"
         Height          =   315
         Left            =   105
         TabIndex        =   34
         ToolTipText     =   "<F10> Insertar Orden de Pedidos"
         Top             =   945
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   12582912
         Text            =   "Producto"
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
      Begin MSDataListLib.DataCombo DCMarca 
         Bindings        =   "FNC.frx":0049
         DataSource      =   "AdoMarca"
         Height          =   285
         Left            =   9870
         TabIndex        =   32
         ToolTipText     =   "<F10> Insertar Orden de Pedidos"
         Top             =   210
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   8388608
         Text            =   "."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCBodega 
         Bindings        =   "FNC.frx":0060
         DataSource      =   "AdoBodega"
         Height          =   315
         Left            =   1155
         TabIndex        =   30
         ToolTipText     =   "<F10> Insertar Orden de Pedidos"
         Top             =   210
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VB.Label Label28 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BODEGA"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   29
         Top             =   210
         Width           =   1065
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   13230
         TabIndex        =   42
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label LabelStockArt 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRODUCTO"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   33
         Top             =   630
         Width           =   9465
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   13230
         TabIndex        =   41
         Top             =   630
         Width           =   1590
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.V.P."
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10500
         TabIndex        =   37
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   9555
         TabIndex        =   35
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESC."
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   11865
         TabIndex        =   39
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label Label29 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Marca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8925
         TabIndex        =   31
         Top             =   210
         Width           =   960
      End
   End
   Begin VB.TextBox TxtDescuento 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   5880
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   52
      Text            =   "FNC.frx":0078
      Top             =   7140
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid DGAsiento_NC 
      Bindings        =   "FNC.frx":007C
      Height          =   2535
      Left            =   105
      TabIndex        =   51
      Top             =   3885
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   8388608
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin VB.TextBox TxtAutorizacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4305
      MaxLength       =   49
      TabIndex        =   23
      Text            =   "0000000000"
      Top             =   1995
      Width           =   6525
   End
   Begin VB.TextBox TextBanco 
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
      Height          =   360
      Left            =   3360
      MaxLength       =   49
      TabIndex        =   7
      Text            =   "."
      Top             =   840
      Width           =   2220
   End
   Begin VB.TextBox TxtIVA 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9765
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "FNC.frx":0098
      Top             =   6825
      Width           =   1905
   End
   Begin VB.TextBox TxtSaldo 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9765
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "FNC.frx":009C
      Top             =   6510
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Nota de Crédito"
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
      Left            =   11760
      Picture         =   "FNC.frx":00A0
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6510
      Width           =   1590
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
      Left            =   13440
      Picture         =   "FNC.frx":04E2
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6510
      Width           =   1590
   End
   Begin VB.TextBox TxtConcepto 
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
      Left            =   3360
      MaxLength       =   110
      TabIndex        =   15
      Top             =   1260
      Width           =   11670
   End
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   105
      Top             =   7350
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoContraCta 
      Height          =   330
      Left            =   1890
      Top             =   7035
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "ContraCta"
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
   Begin MSAdodcLib.Adodc AdoAsiento_NC 
      Height          =   330
      Left            =   1890
      Top             =   7350
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Asiento_NC"
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
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1050
      TabIndex        =   1
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
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "FNC.frx":0DAC
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   105
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
      Top             =   6090
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "FNC.frx":0DC6
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Clientes"
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
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "FNC.frx":0DDD
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   1995
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "000000000"
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
   Begin MSDataListLib.DataCombo DCTC 
      Bindings        =   "FNC.frx":0DF6
      DataSource      =   "AdoTC"
      Height          =   315
      Left            =   105
      TabIndex        =   17
      Top             =   1995
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "FA/NV"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "FNC.frx":0E0A
      DataSource      =   "AdoSerie"
      Height          =   315
      Left            =   1155
      TabIndex        =   19
      Top             =   1995
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "001001"
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
   Begin MSAdodcLib.Adodc AdoTC 
      Height          =   330
      Left            =   105
      Top             =   6720
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "TC"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   105
      Top             =   6405
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   1890
      Top             =   6720
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Left            =   105
      Top             =   7665
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Left            =   1890
      Top             =   6405
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   330
      Left            =   105
      Top             =   7035
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Marca"
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
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contra Cuenta a aplicar la Nota de Credito"
      Height          =   330
      Left            =   7560
      TabIndex        =   12
      Top             =   525
      Width           =   7470
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comp. No. "
      Height          =   330
      Left            =   6405
      TabIndex        =   10
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      Height          =   330
      Left            =   5565
      TabIndex        =   8
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label23 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SubTotal Con IVA"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3990
      TabIndex        =   56
      Top             =   6825
      Width           =   1905
   End
   Begin VB.Label Label22 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SubTotal Sin IVA"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3990
      TabIndex        =   54
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label Label21 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Descuento"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3990
      TabIndex        =   53
      Top             =   7140
      Width           =   1905
   End
   Begin VB.Label LblTotalDC 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9765
      TabIndex        =   48
      Top             =   7140
      Width           =   1905
   End
   Begin VB.Label LblSaldo 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   13020
      TabIndex        =   27
      Top             =   1995
      Width           =   2010
   End
   Begin VB.Label LblTotal 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   10920
      TabIndex        =   25
      Top             =   1995
      Width           =   2010
   End
   Begin VB.Label Label17 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo de la Factura"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   13020
      TabIndex        =   26
      Top             =   1680
      Width           =   2010
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Factura"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   10920
      TabIndex        =   24
      Top             =   1680
      Width           =   2010
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No."
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   20
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " T.D."
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1155
      TabIndex        =   18
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label15 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización del Documento"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4305
      TabIndex        =   22
      Top             =   1680
      Width           =   6525
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Linea de Nota de Credito:"
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   3270
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización Nota de Credito"
      Height          =   330
      Left            =   3360
      TabIndex        =   6
      Top             =   525
      Width           =   2220
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha NC"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Nota de Crebito"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7875
      TabIndex        =   47
      Top             =   7140
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total del I.V.A"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7875
      TabIndex        =   45
      Top             =   6825
      Width           =   1905
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SubTotal"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7875
      TabIndex        =   43
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
      Height          =   330
      Left            =   2415
      TabIndex        =   2
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Motivo de la Nota de Credito"
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   1260
      Width           =   3270
   End
End
Attribute VB_Name = "FNotasDeCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ReIngNC   As Boolean
Dim DocConInv As Boolean
Dim Ok_Inv    As Boolean

Dim Idx  As Long
Dim NCNo As Long



Dim AdoAuxDB As ADODB.Recordset

Private Sub Command2_Click()
Dim SubTotalCosto As Currency
Dim Grupo As String
    FechaValida MBoxFecha
   'MsgBox CCur(LblTotalDC.Caption) & vbCrLf & CCur(LblSaldo.Caption)
    If CCur(LblTotalDC.Caption) <= CCur(LblSaldo.Caption) Then
       If Not ReIngNC Then FA.Nota_Credito = ReadSetDataNum("NC_SERIE_" & FA.Serie_NC, True, True)
        FA.Fecha_NC = MBoxFecha
        Contra_Cta = SinEspaciosIzq(DCContraCta)
        If Len(Contra_Cta) <= 1 Then Contra_Cta = ReadAdoCta("Cta_Devolucion_Ventas")
        Listar_Articulos_Malla
        
        Actualiza_Procesado_Kardex_Factura FA
        
        sSQL = "DELETE * " _
             & "FROM Detalle_Nota_Credito " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Serie = '" & FA.Serie_NC & "' " _
             & "AND Secuencial = " & FA.Nota_Credito & " "
        Ejecutar_SQL_SP sSQL
        
        FA.ClaveAcceso_NC = Ninguno
        FA.SubTotal_NC = 0
        FA.Total_IVA_NC = 0
        FA.Descuento_NC = 0
        Cantidad = 0
        If Len(FA.Autorizacion_NC) >= 13 Then TMail.TipoDeEnvio = "CE"
        With AdoAsiento_NC.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
             Do While Not .EOF
                FA.SubTotal_NC = FA.SubTotal_NC + .fields("SUBTOTAL")
                FA.Total_IVA_NC = FA.Total_IVA_NC + .fields("TOTAL_IVA")
                FA.Descuento_NC = FA.Descuento_NC + .fields("DESCUENTO")
                SubTotalCosto = Redondear(.fields("SUBTOTAL") / .fields("CANT"), 6)
               'SubTotal = Redondear(.Fields("CANT") * SubTotalCosto, 2)
                SubTotal = Redondear(.fields("CANT") * .fields("COSTO"), 2)
                
               'Grabamos el detalle de la NC
               'Cta_Devolucion, , Porc_IVA,
                SetAdoAddNew "Detalle_Nota_Credito"
                SetAdoFields "T", Normal
                SetAdoFields "CodigoC", .fields("Codigo_C")
                SetAdoFields "Cta_Devolucion", Contra_Cta
                SetAdoFields "Fecha", FA.Fecha_NC
                SetAdoFields "Serie", FA.Serie_NC
                SetAdoFields "Secuencial", FA.Nota_Credito
                SetAdoFields "Autorizacion", FA.Autorizacion_NC
                SetAdoFields "Codigo_Inv", .fields("CODIGO")
                SetAdoFields "Cantidad", .fields("CANT")
                SetAdoFields "Producto", .fields("PRODUCTO")
                SetAdoFields "CodBodega", .fields("CodBod")
                SetAdoFields "Total_IVA", .fields("TOTAL_IVA")
                SetAdoFields "Precio", .fields("PVP")
                SetAdoFields "Total", .fields("SUBTOTAL")
                SetAdoFields "CodMar", .fields("CodMar")
                SetAdoFields "Cod_Ejec", .fields("Cod_Ejec")
                SetAdoFields "Porc_C", .fields("Porc_C")
                SetAdoFields "Porc_IVA", .fields("Porc_IVA")
                SetAdoFields "Mes_No", .fields("Mes_No")
                SetAdoFields "Mes", .fields("Mes")
                SetAdoFields "Anio", .fields("Anio")
                SetAdoFields "TC", FA.TC
                SetAdoFields "Serie_FA", FA.Serie
                SetAdoFields "Factura", FA.Factura
                SetAdoFields "A_No", CByte(Ln_No)
                SetAdoUpdate
                
               'Grabamos en el Kardex la factura
                If .fields("Ok") Then
                    SetAdoAddNew "Trans_Kardex"
                    SetAdoFields "T", Normal
                    SetAdoFields "TP", Ninguno
                    SetAdoFields "Numero", 0
                    SetAdoFields "TC", FA.TC
                    SetAdoFields "Serie", FA.Serie
                    SetAdoFields "Fecha", FA.Fecha_NC
                    SetAdoFields "Factura", FA.Factura
                    SetAdoFields "Codigo_P", FA.CodigoC
                    SetAdoFields "CodigoL", FA.Cod_CxC
                    SetAdoFields "CodBodega", .fields("CodBod")
                    SetAdoFields "CodMarca", .fields("CodMar")
                    SetAdoFields "Codigo_Inv", .fields("CODIGO")
                    SetAdoFields "Total_IVA", .fields("Total_IVA")
                    SetAdoFields "Entrada", .fields("CANT")
                    SetAdoFields "PVP", .fields("PVP") 'SubTotalCosto
                    SetAdoFields "Valor_Unitario", .fields("COSTO") 'SubTotalCosto
                    SetAdoFields "Costo", .fields("COSTO")
                    SetAdoFields "Valor_Total", Redondear(.fields("CANT") * .fields("COSTO"), 2)
                    SetAdoFields "Total", Redondear(.fields("CANT") * .fields("COSTO"), 2)
                    SetAdoFields "Descuento", .fields("DESCUENTO")
                    SetAdoFields "Detalle", "NC: " + FA.Serie_NC + "-" + Format(FA.Nota_Credito, "000000000") + " -" + MidStrg(FA.Cliente, 1, 79)
                    SetAdoFields "Cta_Inv", .fields("Cta_Inventario")
                    SetAdoFields "Contra_Cta", .fields("Cta_Costo")
                    SetAdoFields "Item", NumEmpresa
                    SetAdoFields "Periodo", Periodo_Contable
                    SetAdoFields "CodigoU", CodigoUsuario
                    SetAdoUpdate
                   'MsgBox "Grabado"
                End If
               .MoveNext
             Loop
         End If
        End With
        
        TA.T = Normal
        TA.TP = FA.TC
        TA.Serie = FA.Serie
        TA.Factura = FA.Factura
        TA.Autorizacion = FA.Autorizacion
        TA.Fecha = MBoxFecha
        TA.CodigoC = FA.CodigoC
        TA.Cta_CxP = FA.Cta_CxP
        TA.Cta = Contra_Cta
        
        TA.Serie_NC = FA.Serie_NC
        TA.Autorizacion_NC = FA.Autorizacion_NC
        TA.Nota_Credito = FA.Nota_Credito

        TA.Banco = "NOTA DE CREDITO"
        TA.Cheque = "VENTAS SIN IVA"
        TA.Abono = Total_Sin_IVA - Total_Desc
        Grabar_Abonos TA
        
        TA.Banco = "NOTA DE CREDITO"
        TA.Cheque = "VENTAS CON IVA"
        TA.Abono = Total_Con_IVA - Total_Desc2
        Grabar_Abonos TA
        
        TA.Cta = Cta_IVA
        TA.Banco = "NOTA DE CREDITO"
        TA.Cheque = "I.V.A."
        TA.Abono = FA.Total_IVA_NC
        Grabar_Abonos TA

        If TxtConcepto = "" Then TxtConcepto = Ninguno
        
        sSQL = "UPDATE Facturas " _
             & "SET Nota = '" & TxtConcepto & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Factura = " & FA.Factura & " " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' "
        Ejecutar_SQL_SP sSQL
            
        sSQL = "UPDATE Trans_Abonos " _
             & "SET Serie_NC = '" & FA.Serie_NC & "', " _
             & "Autorizacion_NC = '" & FA.Autorizacion_NC & "', " _
             & "Secuencial_NC = '" & FA.Nota_Credito & "', " _
             & "Clave_Acceso_NC = '" & Ninguno & "', " _
             & "Estado_SRI_NC = 'CG' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Factura = " & FA.Factura & " " _
             & "AND TP = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' "
        Ejecutar_SQL_SP sSQL
        If ((FA.SubTotal_NC + FA.Total_IVA_NC) > 0) And Len(FA.Autorizacion_NC) >= 13 Then SRI_Crear_Clave_Acceso_Nota_Credito FA, True
        
    ''''  If SaldoPendiente + SubTotal_IVA > 0 Then
    ''''     Mensajes = "Esta seguro que desea proceder," & vbCrLf _
    ''''              & "con la Nota de Credito"
    ''''     Titulo = "FORMULARIO DE NC"
    ''''     If BoxMensaje = vbYes Then
    ''''        RatonReloj
    ''''        sSQL = "SELECT * " _
    ''''             & "FROM Catalogo_CxCxP " _
    ''''             & "WHERE Item = '" & NumEmpresa & "' " _
    ''''             & "AND Periodo = '" & Periodo_Contable & "' " _
    ''''             & "AND Codigo = '" & CodigoCliente & "' " _
    ''''             & "AND Cta = '" & TA.Cta_CxP & "' " _
    ''''             & "AND TC = 'P' "
    ''''        Select_Adodc AdoComision, sSQL
    ''''        With AdoComision.Recordset
    ''''         If .RecordCount <= 0 Then
    ''''             SetAddNew AdoComision
    ''''             SetFields AdoComision, "Item", NumEmpresa
    ''''             SetFields AdoComision, "Periodo", Periodo_Contable
    ''''             SetFields AdoComision, "Codigo", CodigoCliente
    ''''             SetFields AdoComision, "Cta", TA.Cta_CxP
    ''''             SetFields AdoComision, "TC", "P"
    ''''             SetUpdate AdoComision
    ''''         End If
    ''''        End With
        Ln_No = 0
        sSQL = "DELETE * " _
             & "FROM Asiento_NC " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Ejecutar_SQL_SP sSQL
        
        Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
        
        Listar_Facturas_Pendientes_NC

        Listar_Articulos_Malla
        RatonNormal
        MsgBox "Proceso Terminado con éxito"
        MBoxFecha.SetFocus
    Else
        RatonNormal
        MsgBox "No se puede proceder, El Saldo Pendiente es menor que el total de la Nota de Credito"
        TxtAutorizacion.SetFocus
    End If
End Sub

Private Sub Command3_Click()
  Unload FNotasDeCredito
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
    If KeyCode = vbKeyEscape Then Command2.SetFocus
End Sub

Private Sub DCArticulo_LostFocus()
    Codigos = Ninguno
    BanIVA = False
    Precio = 0
    Producto = DCArticulo
    With AdoArticulo.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Producto = '" & Producto & "' ")
         If Not .EOF Then
            Codigos = .fields("Codigo_Inv")
            BanIVA = .fields("IVA")
            Precio = .fields("PVP")
            Ok_Inv = Leer_Codigo_Inv(Codigos, MBoxFecha)
            CodigoA = InputBox("Detalle del Producto", "DETALLE PRODUCTO NC", Producto)
            If CodigoA <> Producto And Len(CodigoA) > 1 Then Producto = CodigoA
         End If
     End If
    End With
    TextVUnit = Precio
    Listar_Articulos_Malla
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBodega_LostFocus()
    Cod_Bodega = Ninguno
    With AdoBodega.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Bodega Like '" & DCBodega & "' ")
         If Not .EOF Then Cod_Bodega = .fields("CodBod")
     End If
    End With
End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCClientes_LostFocus()
  FA.CodigoC = Ninguno
  FA.Cta_CxP = Ninguno
  FA.Cliente = Ninguno
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCClientes & "' ")
       If Not .EOF Then
          FA.CodigoC = .fields("Codigo")
          FA.Cliente = .fields("Cliente")
       End If
   End If
  End With
  
  sSQL = "SELECT Codigo, Concepto, CxC " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Fact = 'NC' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha <= #" & BuscarFecha(MBoxFecha) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(MBoxFecha) & "# "
  If Len(FA.Cta_CxP) > 2 Then sSQL = sSQL & "AND '" & FA.Cta_CxP & "' IN (CxC,CxC_Anterior) "
  sSQL = sSQL & "ORDER BY CxC, Concepto "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
  'MsgBox AdoLinea.Recordset.RecordCount
  'MsgBox sSQL
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoC = '" & FA.CodigoC & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND TC <> 'OP' " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDB_Combo DCTC, AdoTC, sSQL, "TC"
  If AdoTC.Recordset.RecordCount <= 0 Then MsgBox "Este Cliente no ha empezado a generar facturas"
  TxtConcepto = "Nota de Crédito de: " & DCClientes.Text
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  FA.Factura = Val(DCFactura)
  sSQL = "SELECT T,Fecha,Cta_CxP,Cod_CxC,Porc_IVA,Total_MN,Saldo_MN,IVA,Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND Serie = '" & FA.Serie & "' " _
       & "AND Factura = " & FA.Factura & " " _
       & "AND T <> '" & Anulado & "' " _
       & "AND Saldo_MN > 0 " _
       & "ORDER BY Autorizacion "
  Select_AdoDB AdoAuxDB, sSQL
  With AdoAuxDB
   If .RecordCount > 0 Then
       FA.T = .fields("T")
       FA.Fecha = .fields("Fecha")
       FA.Cta_CxP = .fields("Cta_CxP")
       FA.Cod_CxC = .fields("Cod_CxC")
       FA.Porc_IVA = .fields("Porc_IVA")
       FA.Total_MN = .fields("Total_MN")
       FA.Saldo_MN = .fields("Saldo_MN")
       FA.Autorizacion = .fields("Autorizacion")
       If .fields("IVA") > 0 Then FA.Porc_NC = .fields("Porc_IVA")
       TxtAutorizacion = .fields("Autorizacion")
   End If
  End With
  AdoAuxDB.Close
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  ReIngNC = False
  FA.TC = "NC"
  FA.Cod_CxC = DCLinea
  Lineas_De_CxC FA
  TextBanco = FA.Autorizacion
  TextCheqNo = FA.Serie_NC
 'NC.Factura = Numero_Factura(NC)
  FA.Nota_Credito = ReadSetDataNum("NC_SERIE_" & FA.Serie_NC, True, False)
  TextCompRet = Format(FA.Nota_Credito, "000000000")
  NCNo = FA.Nota_Credito
End Sub

Private Sub DCMarca_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCMarca_LostFocus()
  Cod_Marca = Ninguno
  With AdoMarca.Recordset
   If .RecordCount > 0 Then
      .Find ("Marca = '" & DCMarca & "' ")
       If Not .EOF Then Cod_Marca = .fields("CodMar")
   End If
  End With
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
   FA.Serie = DCSerie
   sSQL = "SELECT Factura " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoC = '" & FA.CodigoC & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND Serie = '" & FA.Serie & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND Saldo_MN > 0 " _
       & "GROUP BY Factura " _
       & "ORDER BY Factura "
  SelectDB_Combo DCFactura, AdoFactura, sSQL, "Factura"
End Sub

Private Sub DCTC_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCTC_LostFocus()
  FA.TC = DCTC
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoC = '" & FA.CodigoC & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub DGAsiento_NC_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el Item " & vbCrLf & "(" _
           & AdoAsiento_NC.Recordset.fields("CODIGO") & ") " _
           & AdoAsiento_NC.Recordset.fields("PRODUCTO") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = vbYes Then Cancel = False Else Cancel = True
End Sub

Private Sub DGAsiento_NC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     Listar_Articulos_Malla
     MsgBox "Proceda a grabar"
     Command2.SetFocus
  End If
  If KeyCode = vbKeyReturn Then
     If AdoAsiento_NC.Recordset.RecordCount > 0 Then
        AdoAsiento_NC.Recordset.MoveNext
        If AdoAsiento_NC.Recordset.EOF Then AdoAsiento_NC.Recordset.MoveFirst
     End If
  End If
End Sub

Private Sub Form_Activate()
  SubTotal_IVA = 0
  SaldoPendiente = 0
  
  sSQL = "DELETE * " _
       & "FROM Asiento_NC " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
    
  sSQL = "SELECT * " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodBod, Bodega "
  SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
   
  sSQL = "SELECT * " _
       & "FROM Catalogo_Marcas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Marca "
  SelectDB_Combo DCMarca, AdoMarca, sSQL, "Marca"

  
  sSQL = "SELECT (Codigo & Space(5) & Cuenta) As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) IN ('1','2','4','5') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCContraCta, AdoContraCta, sSQL, "NomCuenta"
  
  sSQL = "SELECT Producto, Codigo_Inv, PVP, IVA, Cta_Inventario " _
       & "FROM Catalogo_Productos " _
       & "WHERE TC = 'P' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Producto "
  SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
' & "AND LEN(Cta_Inventario) > 2 "
  Listar_Facturas_Pendientes_NC
  Listar_Articulos_Malla
  
  TxtSaldo = Format$(SaldoPendiente, "#,##0.00")
  TxtIVA = Format$(SubTotal_IVA, "#,##0.00")
  LblTotalDC.Caption = Format$(SaldoPendiente + SubTotal_IVA, "#,##0.00")
End Sub

Private Sub Form_Load()
  CentrarForm FNotasDeCredito
  ConectarAdodc AdoTC
  ConectarAdodc AdoLinea
  ConectarAdodc AdoSerie
  ConectarAdodc AdoFactura
  ConectarAdodc AdoMarca
  ConectarAdodc AdoBodega
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoClientes
  ConectarAdodc AdoContraCta
  ConectarAdodc AdoAsiento_NC
  
  DGAsiento_NC.Top = FrmProductos.Top + FrmProductos.Height + 80
  DGAsiento_NC.Height = AdoClientes.Top - (FrmProductos.Top + FrmProductos.Height) - 100
     
  FA.TC = Ninguno
  FA.Serie = Ninguno
  FA.Factura = 0
  Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TextCompRet_GotFocus()
  MarcarTexto TextCompRet
End Sub

Private Sub TextCompRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCompRet_LostFocus()
   If NCNo <> Val(TextCompRet) Then
      Mensajes = "Desea Reprocesar esta Nota de Credito?"
      Titulo = "Formulario de Aprobacion"
      If BoxMensaje = vbYes Then
         ReIngNC = True
         FA.Nota_Credito = Val(TextCompRet)
         DCContraCta.SetFocus
      End If
   End If
End Sub

Private Sub TextDesc_GotFocus()
    MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
Dim SubTotalIVA As Currency
Dim SubTotalDesc As Currency
Dim InsertarItem As Boolean
    TextoValido TextDesc
    SubTotalDesc = 0
    SubTotalIVA = 0
    InsertarItem = True
    If Val(TextCant) > 0 And Val(TextVUnit) > 0 Then
       SubTotalDesc = Val(TextDesc)
       SubTotal = Redondear(Val(TextCant) * Val(TextVUnit), 2)
       If BanIVA And FA.TC <> "NV" Then SubTotalIVA = Redondear((SubTotal - SubTotalDesc) * FA.Porc_NC, 4)
       Total = SubTotal_NC + SubTotal + IVA_NC + SubTotalIVA - SubTotalDesc - Total_Desc
       'If Total <= CCur(LblTotal) Then
          With AdoAsiento_NC.Recordset
           If .RecordCount > 0 Then
              .MoveFirst
               Do While Not .EOF
                  If Codigos = .fields("CODIGO") Then InsertarItem = False
                 .MoveNext
               Loop
           End If
          End With
          If InsertarItem Then
             SetAdoAddNew "Asiento_NC"
             SetAdoFields "CODIGO", Codigos
             SetAdoFields "CANT", CCur(TextCant)
             SetAdoFields "PRODUCTO", Producto
             SetAdoFields "SUBTOTAL", SubTotal
             SetAdoFields "DESCUENTO", SubTotalDesc
             SetAdoFields "TOTAL_IVA", SubTotalIVA
             SetAdoFields "CodBod", Cod_Bodega
             SetAdoFields "CodMar", Cod_Marca
             SetAdoFields "Codigo_C", FA.CodigoC
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "PVP", Redondear(Val(TextVUnit.Text), Dec_PVP)
             SetAdoFields "COSTO", DatInv.Costo
             SetAdoFields "Mes_No", Month(FA.Fecha)
             SetAdoFields "Mes", MesesLetras(Month(FA.Fecha))
             SetAdoFields "Anio", Year(FA.Fecha)
             SetAdoFields "Porc_IVA", FA.Porc_NC
             If DatInv.Con_Kardex Then
               SetAdoFields "Ok", DatInv.Con_Kardex
               SetAdoFields "Cta_Inventario", DatInv.Cta_Inventario
               SetAdoFields "Cta_Costo", DatInv.Cta_Costo_Venta
             End If
             SetAdoFields "A_No", CByte(Ln_No)
             SetAdoUpdate
             Ln_No = Ln_No + 1
          Else
             MsgBox "Este Producto ya se ingreso"
          End If
       'Else
       '  MsgBox "No se puede insertar mas productos para devolver"
       'End If
    End If
    Listar_Articulos_Malla
    DCArticulo.SetFocus
End Sub

Private Sub TextVUnit_Change()
    SubTotal = 0
    If Val(TextCant) <> 0 And Val(TextVUnit) <> 0 Then SubTotal = Redondear(CCur(TextCant) * CCur(TextVUnit), 2)
    LabelVTotal.Caption = Format(SubTotal, "#,#00.00")
End Sub

Private Sub TextVUnit_GotFocus()
    MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

'& "SET SubTotal_NC = ANC.SUBTOTAL, Total_IVA_NC = ANC.TOTAL_IVA, Fecha_NC = #" & BuscarFecha(TA.Fecha) & "# "
Private Sub TxtAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtAutorizacion_LostFocus()
     DocConInv = False
     Ln_No = 0
     sSQL = "DELETE * " _
          & "FROM Asiento_NC " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Ejecutar_SQL_SP sSQL

     LblSaldo.Caption = Format(FA.Saldo_MN, "#,##0.00")
     LblTotal.Caption = Format(FA.Total_MN, "#,##0.00")

     Ln_No = 0
     SQL2 = "SELECT Codigo,Cantidad,Precio,Producto,Total,Total_Desc,Total_Desc2,Total_IVA,CodBodega,CodMarca,Cod_Ejec,Porc_C,Porc_IVA,Mes_No,Mes,Ticket " _
          & "FROM Detalle_Factura " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Factura = " & FA.Factura & " " _
          & "AND Autorizacion = '" & FA.Autorizacion & "' " _
          & "ORDER BY ID "
     Select_Adodc AdoAsiento_NC, SQL2
     With AdoAsiento_NC.Recordset
      If .RecordCount > 0 Then
          FA.Cod_Ejec = .fields("Cod_Ejec")
          FA.Porc_C = .fields("Porc_C")
          NoMes = .fields("Mes_No")
          MiMes = .fields("Mes")
          Cod_Bodega = .fields("CodBodega")
          Do While Not .EOF
             Ok_Inv = Leer_Codigo_Inv(.fields("Codigo"), MBoxFecha)
             SetAdoAddNew "Asiento_NC"
             SetAdoFields "CODIGO", .fields("Codigo")
             SetAdoFields "CANT", .fields("Cantidad")
             SetAdoFields "PRODUCTO", .fields("Producto")
             SetAdoFields "SUBTOTAL", .fields("Total")
             SetAdoFields "DESCUENTO", .fields("Total_Desc") + .fields("Total_Desc2")
             SetAdoFields "TOTAL_IVA", .fields("Total_IVA")
             SetAdoFields "CodBod", .fields("CodBodega")
             SetAdoFields "CodMar", .fields("CodMarca")
             SetAdoFields "Codigo_C", FA.CodigoC
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "PVP", .fields("Precio")
             SetAdoFields "COSTO", DatInv.Costo
             SetAdoFields "Cod_Ejec", .fields("Cod_Ejec")
             SetAdoFields "Porc_C", .fields("Porc_C")
             SetAdoFields "Porc_IVA", .fields("Porc_IVA")
             SetAdoFields "Mes_No", .fields("Mes_No")
             SetAdoFields "Mes", .fields("Mes")
             SetAdoFields "Anio", .fields("Ticket")
             If DatInv.Con_Kardex Then
               SetAdoFields "Ok", DatInv.Con_Kardex
               SetAdoFields "Cta_Inventario", DatInv.Cta_Inventario
               SetAdoFields "Cta_Costo", DatInv.Cta_Costo_Venta
             End If
             SetAdoFields "A_No", CByte(Ln_No)
             SetAdoUpdate
             Ln_No = Ln_No + 1
             DocConInv = DatInv.Con_Kardex
            .MoveNext
          Loop
      End If
     End With
     Label15.Caption = " Autorización del Documento - Fecha de Emision: " & FA.Fecha
     If FA.Porc_NC > 0 Then Label5.Caption = " Total del I.V.A " & (FA.Porc_NC * 100) & "%"
     Listar_Articulos_Malla
     If DocConInv Then DCBodega.SetFocus Else DGAsiento_NC.SetFocus
End Sub

Private Sub TxtConcepto_GotFocus()
  MarcarTexto TxtConcepto
End Sub

Private Sub TxtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtConcepto_LostFocus()
  TextoValido TxtConcepto
End Sub

Private Sub TxtIVA_GotFocus()
  MarcarTexto TxtIVA
End Sub

Private Sub TxtIVA_LostFocus()
  SubTotal_IVA = Val(CCur(TxtIVA))
  SaldoPendiente = Val(CCur(TxtSaldo))
  TxtSaldo = Format$(SaldoPendiente, "#,##0.00")
  TxtIVA = Format$(SubTotal_IVA, "#,##0.00")
  LblTotalDC.Caption = Format$(SaldoPendiente + SubTotal_IVA, "#,##0.00")
End Sub

Private Sub TxtSaldo_GotFocus()
  MarcarTexto TxtSaldo
End Sub

Private Sub TxtSaldo_LostFocus()
  TextoValido TxtSaldo, True
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
   Validar_Porc_IVA MBoxFecha
   FA.Porc_IVA = Porc_IVA
   FA.Fecha = MBoxFecha
   FA.Fecha_NC = MBoxFecha
End Sub

Public Sub Listar_Articulos_Malla()
   SubTotal_NC = 0
   IVA_NC = 0
   Total_Desc = 0
   Total_Desc2 = 0
   Total_Sin_IVA = 0
   Total_Con_IVA = 0
   sSQL = "SELECT * " _
        & "FROM Asiento_NC " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY A_No "
   SQLDec = "SUBTOTAL 2|TOTAL_IVA 2|TOTAL 2|DESCUENTO 2|COSTO 6|."
   Select_Adodc_Grid DGAsiento_NC, AdoAsiento_NC, sSQL, SQLDec
   With AdoAsiento_NC.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           If .fields("TOTAL_IVA") > 0 Then
               IVA_NC = IVA_NC + .fields("TOTAL_IVA")
               Total_Con_IVA = Total_Con_IVA + .fields("SUBTOTAL")
               Total_Desc2 = Total_Desc2 + .fields("DESCUENTO")
           Else
               Total_Sin_IVA = Total_Sin_IVA + .fields("SUBTOTAL")
               Total_Desc = Total_Desc + .fields("DESCUENTO")
           End If
           SubTotal_NC = SubTotal_NC + .fields("SUBTOTAL")
          .MoveNext
        Loop
    End If
   End With
   TxtSinIVA = Format(Total_Sin_IVA, "#,##0.00")
   TxtConIVA = Format(Total_Con_IVA, "#,##0.00")
   TxtSaldo = Format(SubTotal_NC, "#,##0.00")
   TxtIVA = Format(IVA_NC, "#,##0.00")
   TxtDescuento = Format(Total_Desc + Total_Desc2, "#,##0.00")
   LblTotalDC.Caption = Format(SubTotal_NC + IVA_NC - (Total_Desc + Total_Desc2), "#,##0.00")
End Sub

Public Sub Listar_Facturas_Pendientes_NC()
  sSQL = "SELECT C.Grupo, C.Codigo, C.Cliente, SUM(F.Total_MN) As TotFact " _
       & "FROM Clientes As C, Facturas As F " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT F.TC IN ('DO','OP') " _
       & "AND F.T <> 'A' " _
       & "AND F.Saldo_MN <> 0 " _
       & "AND C.Codigo = F.CodigoC " _
       & "GROUP BY C.Grupo, C.Codigo, C.Cliente " _
       & "ORDER BY C.Cliente "
  SelectDB_Combo DCClientes, AdoClientes, sSQL, "Cliente"
End Sub
