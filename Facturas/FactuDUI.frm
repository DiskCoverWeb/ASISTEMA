VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturasDUI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   7185
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   8610
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "FactuDUI.frx":0000
      Top             =   2520
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   7665
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   40
      Text            =   "FactuDUI.frx":0005
      Top             =   2520
      Width           =   960
   End
   Begin VB.TextBox TxtTonelaje 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   6510
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   38
      Text            =   "FactuDUI.frx":0007
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "NOTA DE SALIDA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1800
      Left            =   3255
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   2220
      Begin MSDataListLib.DataList DLAux 
         Bindings        =   "FactuDUI.frx":0009
         DataSource      =   "AdoAux"
         Height          =   1425
         Left            =   105
         TabIndex        =   64
         Top             =   210
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   2514
         _Version        =   393216
         BackColor       =   12632064
         ForeColor       =   16777215
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
   End
   Begin VB.TextBox TextObs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   8400
      MaxLength       =   50
      TabIndex        =   26
      Top             =   1470
      Width           =   3165
   End
   Begin MSDataListLib.DataCombo DCEjecutivo 
      Bindings        =   "FactuDUI.frx":001E
      DataSource      =   "AdoEjecutivo"
      Height          =   315
      Left            =   1155
      TabIndex        =   17
      Top             =   1155
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
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
   Begin VB.CheckBox CheqEjec 
      Caption         =   "Agente:"
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
      Top             =   1155
      Width           =   1065
   End
   Begin VB.TextBox TxtGuiaRem 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   10080
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "FactuDUI.frx":0039
      Top             =   1785
      Width           =   1485
   End
   Begin VB.TextBox TxtFUE 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   7140
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "FactuDUI.frx":003D
      Top             =   1785
      Width           =   1275
   End
   Begin VB.TextBox TxtDAU 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   4935
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "FactuDUI.frx":003F
      Top             =   1785
      Width           =   1275
   End
   Begin VB.TextBox TxtDeclaracion 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   2205
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "FactuDUI.frx":0043
      Top             =   1785
      Width           =   1800
   End
   Begin VB.TextBox TxtFactCom 
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   5145
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "FactuDUI.frx":0047
      Top             =   1470
      Width           =   2115
   End
   Begin MSDataListLib.DataCombo DCSolicitud 
      Bindings        =   "FactuDUI.frx":004B
      DataSource      =   "AdoSolicitud"
      Height          =   315
      Left            =   1575
      TabIndex        =   22
      Top             =   1470
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.CheckBox Check1 
      Caption         =   "Solicitud No."
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
      TabIndex        =   21
      Top             =   1470
      Width           =   1485
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FactuDUI.frx":0066
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   3150
      TabIndex        =   13
      Top             =   840
      Width           =   4950
      _ExtentX        =   8731
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
   Begin MSDataListLib.DataCombo DCGrupo_No 
      Bindings        =   "FactuDUI.frx":007F
      DataSource      =   "AdoGrupo"
      Height          =   315
      Left            =   945
      TabIndex        =   11
      Top             =   840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Grupo"
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
   Begin VB.TextBox TxtDetalle 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3165
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   45
      Top             =   2835
      Visible         =   0   'False
      Width           =   8520
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "FactuDUI.frx":0096
      Height          =   3165
      Left            =   105
      TabIndex        =   62
      Top             =   2835
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   5583
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin VB.CheckBox CheqCta 
      Caption         =   "Cuenta de Ingreso"
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
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10815
      Picture         =   "FactuDUI.frx":00B0
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6405
      Width           =   750
   End
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "FactuDUI.frx":097A
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   5565
      TabIndex        =   7
      Top             =   420
      Width           =   3795
      _ExtentX        =   6694
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
   Begin VB.TextBox TextDesc 
      Alignment       =   1  'Right Justify
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
      Left            =   3045
      MultiLine       =   -1  'True
      TabIndex        =   51
      Text            =   "FactuDUI.frx":0991
      Top             =   6720
      Width           =   1485
   End
   Begin VB.TextBox TextNota 
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
      Left            =   1365
      MaxLength       =   200
      TabIndex        =   61
      Top             =   5985
      Width           =   10200
   End
   Begin VB.TextBox TextFacturaNo 
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
      Left            =   9975
      TabIndex        =   15
      Text            =   "0"
      Top             =   840
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1680
      TabIndex        =   4
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   9975
      Picture         =   "FactuDUI.frx":0998
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6405
      Width           =   750
   End
   Begin VB.CheckBox Cheq 
      Caption         =   "Factura/ME"
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
      Left            =   9450
      TabIndex        =   8
      Top             =   420
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   4200
      TabIndex        =   6
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
      Left            =   2730
      Top             =   3885
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
   Begin MSAdodcLib.Adodc AdoListFact 
      Height          =   330
      Left            =   7140
      Top             =   3885
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
      Caption         =   "ListFact"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   525
      Top             =   4200
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   2730
      Top             =   4200
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
   Begin MSAdodcLib.Adodc AdoEjecutivo 
      Height          =   330
      Left            =   525
      Top             =   4515
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
      Caption         =   "Ejecutivo"
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
      Left            =   4935
      Top             =   3885
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   525
      Top             =   3885
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
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FactuDUI.frx":0DDA
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   36
      ToolTipText     =   "<F10> Insertar Orden de Pedidos"
      Top             =   2520
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
      Height          =   330
      Left            =   2730
      Top             =   4515
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   7140
      Top             =   4200
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
      Caption         =   "Grupo"
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
   Begin MSDataListLib.DataCombo DCCta 
      Bindings        =   "FactuDUI.frx":0DF4
      DataSource      =   "AdoCta"
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
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
   Begin MSDataListLib.DataCombo DCMod 
      Bindings        =   "FactuDUI.frx":0E09
      DataSource      =   "AdoMod"
      Height          =   315
      Left            =   6405
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   4935
      Top             =   4200
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoMod 
      Height          =   330
      Left            =   7140
      Top             =   4515
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
      Caption         =   "Mod"
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
   Begin MSAdodcLib.Adodc AdoSolicitud 
      Height          =   330
      Left            =   4935
      Top             =   4515
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
      Caption         =   "Solicitud"
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
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   9975
      TabIndex        =   44
      Top             =   2520
      Width           =   1590
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
      Left            =   9975
      TabIndex        =   43
      Top             =   2205
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio Unitario"
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
      Left            =   8610
      TabIndex        =   41
      Top             =   2205
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
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
      Left            =   7665
      TabIndex        =   39
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tonelaje"
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
      TabIndex        =   37
      Top             =   2205
      Width           =   1170
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Mercaderia"
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
      TabIndex        =   25
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Guía de Remisión"
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
      Left            =   8400
      TabIndex        =   33
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FUE No."
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
      TabIndex        =   31
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DAU No."
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
      Left            =   3990
      TabIndex        =   29
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Declaracion Aduanera"
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
      TabIndex        =   27
      Top             =   1785
      Width           =   2115
   End
   Begin VB.Label LblSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9999999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   9975
      TabIndex        =   20
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
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
      Left            =   8400
      TabIndex        =   19
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
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
      Left            =   6615
      TabIndex        =   18
      Top             =   1155
      Width           =   1800
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Factura Comercial"
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
      TabIndex        =   23
      Top             =   1470
      Width           =   1800
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Importador"
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
      Left            =   2100
      TabIndex        =   12
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label38 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " G&rupo"
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
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Venc."
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
      TabIndex        =   5
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label LabelTotal 
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
      Left            =   7455
      TabIndex        =   57
      Top             =   6720
      Width           =   1590
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
      Left            =   7455
      TabIndex        =   56
      Top             =   6405
      Width           =   1590
   End
   Begin VB.Label LabelIVA 
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
      Left            =   5985
      TabIndex        =   55
      Top             =   6720
      Width           =   1485
   End
   Begin VB.Label LabelServ 
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
      Left            =   4515
      TabIndex        =   53
      Top             =   6720
      Width           =   1485
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
      Left            =   5985
      TabIndex        =   54
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio 10%"
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
      TabIndex        =   52
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total &Desc."
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
      Left            =   3045
      TabIndex        =   50
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label LabelConIVA 
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
      Left            =   1575
      TabIndex        =   49
      Top             =   6720
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TRAYECTO:"
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
      TabIndex        =   60
      Top             =   5985
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nota de Venta No."
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
      Left            =   8190
      TabIndex        =   14
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Emisióin"
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
      TabIndex        =   3
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   10920
      TabIndex        =   9
      Top             =   420
      Width           =   645
   End
   Begin VB.Label LabelSubTotal 
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
      Left            =   105
      TabIndex        =   47
      Top             =   6720
      Width           =   1485
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total con IVA"
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
      TabIndex        =   48
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total sin IVA"
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
      TabIndex        =   46
      Top             =   6405
      Width           =   1485
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   35
      Top             =   2205
      Width           =   6420
   End
End
Attribute VB_Name = "FacturasDUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PorCodigo As Boolean

Public Sub Insertar_Solicitud(Si_o_No As Boolean)
  ' MsgBox CCur(TextCant.Text) * CCur(TextVUnit.Text)
    Real1 = Redondear(Cantidad * Precio, 2)
    If Si_o_No Then
       If BanCli Then Real2 = 0 Else Real2 = Redondear(Real1 * Comision, 2)
       If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
    Else
       Real2 = 0
       If BanIVA Then Real3 = Redondear((Real1) * Porc_IVA, 2) Else Real3 = 0
    End If
    If TipoFactura = "NV" Then Real3 = 0
    SetAdoAddNew "Asiento_F"
    SetAdoFields "CODIGO", CodigoInv
    SetAdoFields "CODIGO_L", CodigoL
    SetAdoFields "PRODUCTO", Producto
    SetAdoFields "CANT", Cantidad
    SetAdoFields "PRECIO", Precio
    SetAdoFields "TOTAL", Real1
    SetAdoFields "Total_Desc", Real2
    SetAdoFields "Total_IVA", Real3
    SetAdoFields "Cta", Cta_Ventas
    SetAdoFields "Item", NumEmpresa
    SetAdoFields "CodigoU", CodigoUsuario
    SetAdoUpdate
End Sub

Public Sub ProcGrabar()
 'Seteamos los encabezados para las facturas
  Calculos_Totales_Factura FA
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     TextoValido TextObs
     TextoValido TxtDAU
     TextoValido TxtFUE
     TextoValido TextNota
     TextoValido TxtDeclaracion
     TextoValido TxtGuiaRem
     TextoValido TxtFactCom
     
     'TextoValido TextComision, True
     FechaValida MBoxFechaV
     FechaTexto = MBoxFecha.Text
     Total_FacturaME = 0
     HoraTexto = Format$(Time, FormatoTimes)
     If Check1.value = 1 Then Moneda_US = True Else Moneda_US = False
     Moneda_US = False
     If Moneda_US Then
        Total_Factura = Redondear((Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio) * Dolar, 2)
        Total_FacturaME = Redondear(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
     Else
        Total_Factura = Redondear(Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_IVA + Total_Servicio, 2)
        Total_FacturaME = 0
     End If
     Saldo = Total_Factura
     Saldo_ME = Total_FacturaME
     If Saldo < 0 Then Saldo = 0
     If TipoFactura = "NV" Then
        NumComp = ReadSetDataNum("Nota Ventas", True, False)
     Else
        NumComp = ReadSetDataNum("Facturas", True, False)
     End If
     If Val(TextFacturaNo.Text) <> NumComp Then
        Factura_No = Val(TextFacturaNo.Text)
     Else
        If TipoFactura = "NV" Then
           Factura_No = ReadSetDataNum("Nota Ventas", True, True)
        Else
           Factura_No = ReadSetDataNum("Facturas", True, True)
        End If
     End If
     Control_Procesos "G", "Grabar Factura No. " & Factura_No
     sSQL = "DELETE * " _
          & "FROM Detalle_Factura " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Facturas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Comprobantes " _
          & "WHERE Numero = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = '" & TipoFactura & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Trans_SubCtas " _
          & "WHERE Numero = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = '" & TipoFactura & "' "
     ConectarAdoExecute sSQL
     
     sSQL = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     SelectAdodc AdoAux, sSQL
     TextoFormaPago = PagoCred
     T = Pendiente
    'Grabamos el numero de factura
     TextoProc = Ninguno
     SetAddNew AdoAux
     SetFields AdoAux, "T", Pendiente
     SetFields AdoAux, "ME", Moneda_US
     SetFields AdoAux, "TC", TipoFactura
     SetFields AdoAux, "Factura", Factura_No
     SetFields AdoAux, "Fecha", FechaTexto
     SetFields AdoAux, "Fecha_C", FechaTexto
     SetFields AdoAux, "Fecha_V", MBoxFechaV.Text
     SetFields AdoAux, "CodigoC", CodigoCliente
     SetFields AdoAux, "Cod_CxC", CodigoL
     SetFields AdoAux, "Forma_Pago", TextoFormaPago
     SetFields AdoAux, "Sin_IVA", Redondear(Total_Sin_IVA, 2)
     SetFields AdoAux, "Con_IVA", Redondear(Total_Con_IVA, 2)
     SetFields AdoAux, "SubTotal", Redondear(Total_Sin_IVA + Total_Con_IVA, 2)
     SetFields AdoAux, "Descuento", Redondear(Total_Desc, 2)
     SetFields AdoAux, "IVA", Redondear(Total_IVA, 2)
     SetFields AdoAux, "Servicio", Redondear(Total_Servicio, 2)
     SetFields AdoAux, "Total_MN", 0
     SetFields AdoAux, "Total_ME", 0
     'SetFields AdoAux, "Comision", Redondear((Total_Con_IVA + Total_Sin_IVA) * Val(TextComision.Text) / 100, 2)
     If Moneda_US Then
        SetFields AdoAux, "Total_ME", Total_FacturaME
        Total = Total_FacturaME
     Else
        SetFields AdoAux, "Total_MN", Total_Factura
        Total = Total_Factura
     End If
     SetFields AdoAux, "Saldo_MN", Total_Factura
     SetFields AdoAux, "Saldo_ME", Total_FacturaME
     If CheqEjec.value = 1 Then
        SetFields AdoAux, "Cod_Ejec", CodigoEjecutivo
        'SetFields AdoAux, "Porc_C", Redondear(Val(TextComision.Text) / 100, 4)
     End If
     SetFields AdoAux, "Cotizacion", Dolar
     SetFields AdoAux, "Observacion", TextObs.Text
     'SetFields AdoAux, "Definitivo", TextDef.Text
     SetFields AdoAux, "Nota", TextNota.Text
     SetFields AdoAux, "Cta_CxP", Cta_Cobrar
     SetFields AdoAux, "Cta_Venta", Cta_Ventas
     SetFields AdoAux, "DAU", TxtDAU
     SetFields AdoAux, "FUE", TxtFUE
     SetFields AdoAux, "Declaracion", TxtDeclaracion
     SetFields AdoAux, "Remision", TxtGuiaRem
     SetFields AdoAux, "Comercial", TxtFactCom
     SetFields AdoAux, "Cantidad", NivelNo
     SetFields AdoAux, "Kilos", PLitro
     SetFields AdoAux, "Hora", Format$(Time, FormatoTimes)
     SetFields AdoAux, "Solicitud", Solicitud_No
     SetFields AdoAux, "CodigoU", CodigoUsuario
     SetFields AdoAux, "Item", NumEmpresa
     SetUpdate AdoAux
     Habitacion_No = Ninguno
     sSQL = "SELECT * " _
          & "FROM Detalle_Factura " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     SelectAdodc AdoAux, sSQL
     With AdoAsientoF.Recordset
         .MoveFirst
          Do While Not .EOF
            SetAddNew AdoAux
            SetFields AdoAux, "T", Pendiente
            SetFields AdoAux, "TC", TipoFactura
            SetFields AdoAux, "No_Hab", .Fields("HABIT")
            SetFields AdoAux, "Factura", Factura_No
            SetFields AdoAux, "CodigoC", CodigoCliente
            SetFields AdoAux, "Fecha", FechaTexto
            SetFields AdoAux, "Codigo", .Fields("CODIGO")
            SetFields AdoAux, "Cantidad", .Fields("CANT")
            SetFields AdoAux, "CodigoL", CodigoL
            SetFields AdoAux, "Reposicion", .Fields("REP")
            SetFields AdoAux, "Tonelaje", .Fields("TONELAJE")
            SetFields AdoAux, "Precio", .Fields("PRECIO")
            SetFields AdoAux, "Total", .Fields("TOTAL")
            SetFields AdoAux, "Total_Desc", .Fields("Total_Desc")
            SetFields AdoAux, "Total_IVA", .Fields("Total_IVA")
            SetFields AdoAux, "Producto", .Fields("PRODUCTO")
            SetFields AdoAux, "Cod_Ejec", .Fields("Cod_Ejec")
            SetFields AdoAux, "Porc_C", .Fields("Porc_C")
            'SetFields AdoAux, "Ruta", TextDef.Text
            SetFields AdoAux, "CodigoU", CodigoUsuario
            SetFields AdoAux, "Item", NumEmpresa
            SetUpdate AdoAux
            If .Fields("HABIT") <> Ninguno Then Habitacion_No = .Fields("HABIT")
           .MoveNext
          Loop
     End With
   ' Ingresar a Trans_SubCta
     If CheqCta.value = 1 Then
     sSQL = "SELECT * " _
          & "FROM Trans_SubCtas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     SelectAdodc AdoAux, sSQL
     With AdoAsientoF.Recordset
         .MoveFirst
          Do While Not .EOF
             LeerCta .Fields("Cta")
             If Codigo <> Ninguno And SubCta = "I" Then
                SetAddNew AdoAux
                SetFields AdoAux, "T", Normal
                SetFields AdoAux, "TC", SubCta
                SetFields AdoAux, "Factura", Factura_No
                SetFields AdoAux, "Codigo", CodigoCliente
                SetFields AdoAux, "Fecha", FechaTexto
                SetFields AdoAux, "Fecha_V", MBoxFechaV.Text
                SetFields AdoAux, "TP", TipoFactura
                SetFields AdoAux, "Numero", Factura_No
                SetFields AdoAux, "Cta", .Fields("Cta")
                SetFields AdoAux, "Debitos", .Fields("TOTAL")
                SetFields AdoAux, "Item", NumEmpresa
                SetUpdate AdoAux
             End If
            .MoveNext
          Loop
     End With
   ' Ingresar a Comprobantes
     sSQL = "SELECT * " _
          & "FROM Comprobantes " _
          & "WHERE Numero = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = '" & TipoFactura & "' "
     SelectAdodc AdoAux, sSQL
     With AdoAsientoF.Recordset
         .MoveFirst
          Do While Not .EOF
             LeerCta .Fields("Cta")
             If Codigo <> Ninguno And SubCta = "I" Then
                SetAddNew AdoAux
                SetFields AdoAux, "T", Normal
                SetFields AdoAux, "TP", TipoFactura
                SetFields AdoAux, "Numero", Factura_No
                SetFields AdoAux, "Codigo_B", CodigoCliente
                SetFields AdoAux, "Fecha", FechaTexto
                SetFields AdoAux, "Concepto", "Factura No. " & Format$(Factura_No, "0000000") & ", " & .Fields("PRODUCTO")
                SetFields AdoAux, "Efectivo", .Fields("TOTAL")
                SetFields AdoAux, "Monto_Total", .Fields("TOTAL")
                SetFields AdoAux, "CodigoU", CodigoUsuario
                SetFields AdoAux, "Item", NumEmpresa
                SetUpdate AdoAux
             End If
            .MoveNext
          Loop
     End With
     End If
     sSQL = "DELETE * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Trans_Pedidos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND No_Hab = '" & Habitacion_No & "' "
     ConectarAdoExecute sSQL
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SelectAdodc AdoAsientoF, sSQL
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
    'Grabamos el numero de factura
     RatonNormal
     Bandera = False
     Evaluar = True
     Mensajes = "Pago al Contado"
     Titulo = "Formulario de Grabacion"
     Numero = Factura_No
     If BoxMensaje = vbYes Then Abonos.Show 1
     Factura_No = Numero
     Imprimir_Facturas AdoFactura, AdoAsientoF, FA
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
End Sub

Public Sub DatosArticulos()
  With AdoArticulo.Recordset
       Producto = .Fields("Producto")
       Cta_Ventas = .Fields("Cta_Ventas")
       TextVUnit.Text = Format$(.Fields("PVP"), "#,##0.00")
       'LabelStock.Caption = .Fields("Stock_Actual")
       Codigos = .Fields("Codigo_Inv")
       BanIVA = .Fields("IVA")
       If TipoFactura = "NV" Then BanIVA = False
       DCArticulo.Text = Producto
       'TextComEjec.Text = "0"
       TxtDetalle.Visible = True
       TxtDetalle.SetFocus
       TxtDetalle.Text = Producto
       If Len(.Fields("Detalle")) > 1 Then TxtDetalle.Text = TxtDetalle.Text & vbCrLf & .Fields("Detalle")
  End With
End Sub

Private Sub Check1_Click()
  If Check1.value = 1 Then
     DCSolicitud.Visible = True
  Else
     DCSolicitud.Visible = False
  End If
End Sub

Private Sub CheqCta_Click()
  If CheqCta.value = 1 Then
     DCCta.Visible = True
     DCMod.Visible = True
  Else
     DCCta.Visible = False
     DCMod.Visible = False
  End If
End Sub

Private Sub CheqEjec_Click()
 If CheqEjec.value = 1 Then
    DCEjecutivo.Visible = True
    'Label11.Visible = True
 Else
    DCEjecutivo.Visible = False
    'Label11.Visible = False
 End If
End Sub

Private Sub CheqEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Command1_Click()
  Unload FacturasDUI
End Sub

Private Sub Command2_Click()
  Mensajes = "Esta Seguro que desea grabar: " & vbCrLf & " La Factura No. " & TextFacturaNo.Text
  Titulo = "Formulario de Grabacion"
  If BoxMensaje = vbYes Then
     If Check1.value = 1 Then Moneda_US = True Else Moneda_US = False
     Moneda_US = False
     Calculos_Totales_Factura FA
     TextoFormaPago = PagoCred
     ProcGrabar
     sSQL = "DELETE * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
  End If
  If CheqCta.value = 1 Then
     sSQL = "DELETE * " _
          & "FROM Comprobantes " _
          & "WHERE TP = '" & TipoFactura & "' " _
          & "AND Numero = " & Factura_No & " " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Trans_SubCtas " _
          & "WHERE TP = '" & TipoFactura & "' " _
          & "AND Numero = " & Factura_No & " " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
    'Encabezado Comprobante
     SetAdoAddNew "Comprobantes"
     SetAdoFields "T", Normal
     SetAdoFields "TP", TipoFactura
     SetAdoFields "Numero", Factura_No
     SetAdoFields "Codigo_B", CodigoCliente
     SetAdoFields "Fecha", FechaTexto
     SetAdoFields "Concepto", "Factura No. " & Format$(Factura_No, "0000000") & "."
     SetAdoFields "Efectivo", Total_Factura
     SetAdoFields "Monto_Total", Total_Factura
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoUpdate
    'El SubModulo
     SetAdoAddNew "Trans_SubCtas"
     SetAdoFields "T", Normal
     SetAdoFields "TP", TipoFactura
     SetAdoFields "Numero", Factura_No
     SetAdoFields "Factura", Factura_No
     SetAdoFields "Fecha", FechaTexto
     SetAdoFields "Fecha_V", FechaTexto
     SetAdoFields "Creditos", Total_Factura
     SetAdoFields "TC", "I"
     SetAdoFields "Cta", SinEspaciosDer(DCCta.Text)
     SetAdoFields "Codigo", SinEspaciosDer(DCMod.Text)
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "ID", Factura_No
     SetAdoUpdate
  End If
  If TipoFactura = "NV" Then
     NumComp = ReadSetDataNum("Nota Ventas", True, False)
  Else
     NumComp = ReadSetDataNum("Facturas", True, False)
  End If
  TextFacturaNo.Text = NumComp
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
  If Mod_Fact Then TextFacturaNo.SetFocus Else DCGrupo_No.SetFocus
End Sub

Private Sub DCEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCEjecutivo_LostFocus()
  With AdoEjecutivo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCEjecutivo.Text & "' ")
       If Not .EOF Then
          CodigoEjecutivo = .Fields("Codigo")
          CodigoCorresp = DCEjecutivo.Text
          'Comision = .Fields("Porc_C")
       Else
          MsgBox "Cuenta No Asignada"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCGrupo_No_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupo_No_LostFocus()
  If DCGrupo_No.Text = "" Then DCGrupo_No.Text = Ninguno
  If DCGrupo_No.Text <> Ninguno Then
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Cliente <> '.' " _
          & "AND Grupo = '" & DCGrupo_No.Text & "' " _
          & "AND FA <> " & Val(adFalse) & " " _
          & "ORDER BY Cliente "
  Else
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Cliente <> '.' " _
          & "AND FA <> " & Val(adFalse) & " " _
          & "AND T = 'N' " _
          & "ORDER BY Cliente "
  End If
  SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
  Label9.Caption = "0"
  Grupo_No = DCGrupo_No.Text
  If AdoCliente.Recordset.RecordCount > 0 Then Label9.Caption = AdoCliente.Recordset.RecordCount
End Sub

Private Sub DCLinea_GotFocus()
  Grupo_No = Ninguno
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  FA.Cod_CxC = DCLinea.Text
  Lineas_De_CxC FA
End Sub

Private Sub DCArticulo_GotFocus()
  'LabelStock.Caption = ""
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         Empleados = False
         Calculos_Totales_Factura FA
         Command2.SetFocus
    Case vbKeyF10
  End Select
End Sub

Private Sub DCArticulo_LostFocus()
  Codigos = Ninguno
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & DCArticulo.Text & "' ")
       If Not .EOF Then
          DatosArticulos
       Else
         .MoveFirst
         .Find ("Codigo_Inv Like '" & DCArticulo.Text & "' ")
          If Not .EOF Then
             DatosArticulos
          Else
            .MoveFirst
            .Find ("Codigo_Barra Like '" & DCArticulo.Text & "' ")
             If Not .EOF Then
                DatosArticulos
             Else
               .MoveFirst
               .Find ("Codigo_Inv Like '" & SinEspaciosDer(DCArticulo.Text) & "' ")
                If Not .EOF Then
                   DatosArticulos
                Else
                   MsgBox "Producto no Asignado"
                End If
             End If
          End If
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCCliente_GotFocus()
  MarcarTexto DCCliente
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  LabelCodigo.Caption = "Ninguno"
  CodigoCliente = Ninguno
  NombreCliente = Ninguno
  DireccionCli = Ninguno
  CodigoCorresp = Ninguno
  LabelCodigo.Caption = CodigoCliente
  'LabelTelefono.Caption = Ninguno
  'LabelRUC.Caption = Ninguno
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCCliente.Text & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          NombreCliente = DCCliente.Text
          LabelCodigo.Caption = CodigoCliente
          BanCli = .Fields("Descuento")
          Comision = .Fields("Porc_C")
          'LabelRUC.Caption = .Fields("CI_RUC")
          DireccionCli = .Fields("Direccion")
          If Mod_Fact Then TextFacturaNo.SetFocus Else CheqEjec.SetFocus
       Else
          Nuevo = True
          NombreCliente = DCCliente.Text
          Facturas.Visible = False
          MsgBox "Cliente no Asignado"
          FClientesFlash.Show 1
          Facturas.Visible = True
          DCGrupo_No.SetFocus
       End If
       sSQL = "SELECT * " _
            & "FROM Trans_Aduanas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodigoC = '" & CodigoCliente & "' " _
            & "AND T = 'P' " _
            & "ORDER BY Solicitud_No "
       SelectDBCombo DCSolicitud, AdoSolicitud, sSQL, "Solicitud_No"
   Else
       'MsgBox "No existen datos"
       Nuevo = True
       NombreCliente = DCCliente.Text
       FacturasDUI.Visible = False
       MsgBox "Cliente no Asignado"
       FClientesFlash.Show 1
       FacturasDUI.Visible = True
       DCGrupo_No.SetFocus
   End If
  End With
  sSQL = "SELECT CodigoC,SUM(Saldo_MN) As Saldo_Pend " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoC = '" & CodigoCliente & "' " _
       & "AND Saldo_MN > 0 " _
       & "GROUP BY CodigoC "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     LblSaldo.Caption = Format$(AdoAux.Recordset.Fields("Saldo_Pend"), "#,##0.00")
  Else
     LblSaldo.Caption = "0.00"
  End If
End Sub

Private Sub DCSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSolicitud_LostFocus()
 PLitro = 0
 CantidadAnt = 0
 With AdoSolicitud.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Solicitud_No Like " & DCSolicitud.Text & " ")
       If Not .EOF Then
          If BanCli = False Then Codigo1 = DCEjecutivo.Text Else Codigo1 = ""
          TextObs.Text = .Fields("Mercaderia")
          TxtDAU.Text = .Fields("DUI")
          PLitro = .Fields("Peso_Bruto")
          NivelNo = .Fields("Cantidad") & " " & .Fields("Clase")
          sSQL = "SELECT * " _
               & "FROM Trans_Kardex " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Solicitud = " & Val(DCSolicitud.Text) & " " _
               & "AND Salida > 0 " _
               & "ORDER BY Fecha,Numero "
          SelectDBList DLAux, AdoAux, sSQL, "Numero"
          Solicitud_No = Val(DCSolicitud.Text)
          Frame1.Visible = True
          DLAux.SetFocus
       Else
          MsgBox "Solicitud No Asignada"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DGAsientoF_AfterDelete()
  'Calculos_Totales_Factura Facturas, AdoAsientoF
End Sub

Private Sub DGAsientoF_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & Chr(13) & "(" _
           & AdoAsientoF.Recordset.Fields("CODIGO") & ") " _
           & AdoAsientoF.Recordset.Fields("PRODUCTO") & "   TOTAL -> " _
           & AdoAsientoF.Recordset.Fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = vbYes Then Cancel = False Else Cancel = True
End Sub

Private Sub DGAsientoF_DblClick()
  TxtDetalle.Visible = False
  TxtDetalle.Text = ""
  With AdoArticulo.Recordset
   If .RecordCount Then
       Codigo4 = DGAsientoF.Columns(0)
      .MoveFirst
      .Find ("Codigo_Inv = '" & Codigo4 & "' ")
       If Not .EOF And .Fields("Detalle") <> Ninguno Then
          TxtDetalle.Visible = True
          TxtDetalle.Text = DGAsientoF.Columns(1) & ": " & vbCrLf & .Fields("Detalle")
          TxtDetalle.SetFocus
       End If
   End If
  End With
End Sub

Private Sub DGAsientoF_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then TxtDetalle.Visible = False
End Sub

Private Sub DLAux_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLAux_LostFocus()
  TxtDetalle.Text = ""
  If AdoAux.Recordset.RecordCount > 0 Then
     Comp_No = DLAux.Text
     sSQL = "DELETE * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
  
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
          
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Solicitud = " & Val(DCSolicitud.Text) & " " _
          & "AND Entrada > 0 " _
          & "ORDER BY Fecha,Numero "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          FechaInicial = .Fields("Fecha")
      End If
     End With
     Codigo1 = Ninguno
     Solicitud_No = Val(DCSolicitud.Text)
     TxtTonelaje.Text = PLitro
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Solicitud = " & Val(DCSolicitud.Text) & " " _
          & "AND Numero = " & Comp_No & " " _
          & "AND Salida > 0 " _
          & "ORDER BY Fecha,Numero "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          FechaFinal = .Fields("Fecha")
          Cuenta_No = .Fields("DUI")
          Cantidad = .Fields("Salida")
          Codigo1 = .Fields("Cod_Tarifa")
          CantidadAnt = .Fields("Salida")
      End If
     End With
     sSQL = "SELECT * " _
          & "FROM Catalogo_Tarifas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo = '" & Codigo1 & "' "
     SelectAdodc AdoMod, sSQL
     With AdoMod.Recordset
      If .RecordCount > 0 Then
          NoDias = CFechaLong(FechaFinal) - CFechaLong(FechaInicial)
'''          TxtDetalle.Text = "DUI No.          :  " & Cuenta_No & vbCrLf _
'''                          & "FECHA DE INGRESO :  " & FechaInicial & vbCrLf _
'''                          & "FECHA DE SALIDA  :  " & FechaFinal & vbCrLf _
'''                          & "No. DIAS         :  " & NoDias & vbCrLf _
'''                          & " " & vbCrLf _
'''                          & " " & vbCrLf _
'''                          & " " & vbCrLf _
'''                          & " " & CodigoCorresp & vbCrLf _
'''                          & " " & vbCrLf _
'''                          & " " & vbCrLf _
'''                          & "                    ALMACENAJE"
          TxtDetalle.Text = "                      " & Cuenta_No & vbCrLf & vbCrLf _
                          & "                      " & FechaInicial & vbCrLf & vbCrLf _
                          & "                      " & FechaFinal & vbCrLf & vbCrLf _
                          & "                      " & NoDias & vbCrLf _
                          & " " & vbCrLf _
                          & " " & vbCrLf _
                          & " " & CodigoCorresp & vbCrLf _
                          & " " & vbCrLf _
                          & " " & vbCrLf _
                          & "                      ALMACENAJE"
         'Calculo Automatico:
          'MsgBox .Fields("Tarifa")
          BanCli = .Fields("Regalias")
          Comision = .Fields("Comision")
          NoDias = NoDias - .Fields("Desc_Dia")
          If .Fields("Por_Sem") Then
              Total = (.Fields("Tarifa") * (NoDias / 7)) * Cantidad
          ElseIf .Fields("Por_Dia") Then
              Total = .Fields("Tarifa") * NoDias * Cantidad
          End If
          If Total > .Fields("CIF") Then
             If .Fields("Por_Sem") Then
                 Total = ((.Fields("Tarifa") - (.Fields("Tarifa") * .Fields("Descuento"))) * (NoDias / 7)) * Cantidad
             ElseIf .Fields("Por_Dia") Then
                 Total = (.Fields("Tarifa") - (.Fields("Tarifa") * .Fields("Descuento"))) * NoDias * Cantidad
             End If
          End If
      End If
     End With
  End If
  LabelStockArt.Caption = "PRODUCTO " & Space(65) & "Regalia: " & Format$(BanCli, "Yes/No") & " - " & Format$(Comision, "##.00%")
  TextVUnit.Text = Total
  Frame1.Visible = False
  TxtDetalle.Visible = True
  TxtDetalle.SetFocus
End Sub

Private Sub Form_Activate()
  TextFacturaNo.Enabled = Mod_Fact
  CantFact = 1
  PorCodigo = ReadSetDataNum("PorCodigo", True, False)
  NumFacturas = ReadSetDataNum("No_FacturasImp", True, False)
  'Orientacion_Pagina = ReadSetDataNum("OrientacionFact", True, False)
   If TipoFactura = "NV" Then
      FacturasDUI.Caption = "INGRESAR NOTA DE VENTA"
      Label2.Caption = " NOTA DE VENTA No."
      NumComp = ReadSetDataNum("Nota Ventas", True, False)
      Label3.Caption = "I.V.A. 0.00%"
   Else
      FacturasDUI.Caption = "INGRESAR FACTURA"
      Label2.Caption = " FACTURA No."
      NumComp = ReadSetDataNum("Facturas", True, False)
      Label3.Caption = "I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
   End If
   FacturasDUI.Caption = FacturasDUI.Caption & " (" & TipoFactura & ")"
   TextFacturaNo.Text = NumComp
   'TextGrupo_No.Text = ""
   Label36.Caption = "Serv. " & Format$(Porc_Serv * 100, "#0.00") & "%"
   TextCant.Text = "0"
   TextVUnit.Text = "0"
   LabelVTotal.Caption = "0"
   Modificar = False
   Bandera = True
   Mifecha = BuscarFecha(FechaSistema)
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE TL <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Codigo "
   SelectDBCombo DCLinea, AdoLinea, sSQL, "Concepto"
   
   Listar_Productos DCArticulo, AdoArticulo
   sSQL = "SELECT * " _
        & "FROM Clientes " _
        & "WHERE Cliente <> '.' " _
        & "AND FA <> " & Val(adFalse) & " " _
        & "AND T = 'N' " _
        & "ORDER BY Cliente "
   SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
   sSQL = "SELECT Grupo " _
        & "FROM Clientes " _
        & "WHERE T = 'N' " _
        & "AND FA <> " & Val(adFalse) & " " _
        & "GROUP BY Grupo " _
        & "ORDER BY Grupo "
   SelectDBCombo DCGrupo_No, AdoGrupo, sSQL, "Grupo"
   
   sSQL = "SELECT Cuenta & '    ' & Codigo As NomCuenta " _
        & "FROM Catalogo_Cuentas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TC = 'I' " _
        & "AND DG = 'D' " _
        & "ORDER BY Cuenta "
   SelectDBCombo DCCta, AdoCta, sSQL, "NomCuenta"
   
   sSQL = "SELECT Detalle & '    ' & Codigo As NomDetalle " _
        & "FROM Catalogo_SubCtas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TC = 'I' " _
        & "ORDER BY Detalle "
   SelectDBCombo DCMod, AdoMod, sSQL, "NomDetalle"

   sSQL = "SELECT P.Codigo,C.Cliente " _
       & "FROM Clientes As C,Catalogo_CxCxP As P " _
       & "WHERE P.TC = 'P' " _
       & "AND P.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = P.Codigo " _
       & "GROUP BY P.Codigo,C.Cliente " _
       & "ORDER BY C.Cliente "
   SelectDBCombo DCEjecutivo, AdoEjecutivo, sSQL, "Cliente"
  
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
   
   sSQL = "SELECT * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SelectDataGrid DGAsientoF, AdoAsientoF, sSQL
   Label9.Caption = "0"
   If AdoCliente.Recordset.RecordCount > 0 Then Label9.Caption = AdoCliente.Recordset.RecordCount
   RatonNormal
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FacturasDUI
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoMod
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoLinea
   ConectarAdodc AdoCliente
   ConectarAdodc AdoFactura
   ConectarAdodc AdoAsientoF
   ConectarAdodc AdoSolicitud
   ConectarAdodc AdoListFact
   ConectarAdodc AdoArticulo
   ConectarAdodc AdoEjecutivo
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
   FechaTexto1 = MBoxFecha.Text
   PLitro = 0
   CantidadAnt = 0
   Solicitud_No = 0
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV
End Sub

Private Sub TextCant_Change()
   Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
   LabelVTotal.Caption = Format$(Real1, "#,##0.00")
End Sub

Private Sub TextCant_GotFocus()
  TextCant.Text = Cantidad
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  If TextCant.Text = "" Then TextCant.Text = "0"
  Real1 = Val(TextCant.Text) * Val(TextVUnit.Text)
  LabelVTotal.Caption = Format$(Real1, "#,##0.00")
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
  TextoValido TextDesc, True
  Calculos_Totales_Factura FA
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
 If TextFacturaNo.Text = "" Then TextFacturaNo.Text = "0"
 Factura_No = Val(TextFacturaNo.Text)
 sSQL = "SELECT * " _
      & "FROM Facturas " _
      & "WHERE Factura = " & Factura_No & " " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = '" & TipoFactura & "' "
 SelectAdodc AdoFactura, sSQL
 If AdoFactura.Recordset.RecordCount > 0 Then
    MsgBox "Warning:" & vbCrLf _
           & "Ya existe la Factura No. " & Format$(Factura_No, "000000") & "."
    'DCGrupo_No.SetFocus
 End If
End Sub

Private Sub TextNota_GotFocus()
   MarcarTexto TextNota
End Sub

Private Sub TextNota_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNota_LostFocus()
  TextoValido TextNota
End Sub

Private Sub TextObs_GotFocus()
  MarcarTexto TextObs
End Sub

Private Sub TextObs_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
''  If KeyCode = vbKeyF2 Then
''     Frame1.Visible = True
''     sSQL = "SELECT Factura " _
''          & "FROM Facturas " _
''          & "WHERE Codigo_C = '" & CodigoCliente & "' " _
''          & "ORDER BY Codigo_C "
''     SelectDBList DBLFact, AdoListFact, sSQL, "Factura"
''  End If
End Sub

Private Sub TextObs_LostFocus()
  TextoValido TextObs
End Sub

Private Sub TextVUnit_Change()
   LabelVTotal.Caption = Format$(Real1, "#,##0.000")
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
    If Comision > 0 Then
       BanIVA = True
       CodigoInv = "99.01"
       Cantidad = Val(TextCant)
       Precio = Val(TextVUnit)
       Producto = TxtDetalle.Text
       Insertar_Solicitud True
       Calculos_Totales_Factura FA
       Cantidad = 1
       BanIVA = False
       Precio = Redondear((Total_Sin_IVA + Total_Con_IVA - Total_Desc + Total_Servicio) * Comision, 2)
       CodigoInv = "99.02"
       Producto = "2% DE REGALIAS-CAE"
       If BanCli Then Insertar_Solicitud False
       Calculos_Totales_Factura FA
   Else
      TextoValido TextVUnit, True
      TextoValido TextCant, True
   'TextoValido TextDesc1, True
   Real1 = 0: Real2 = 0: Real3 = 0
   With AdoAsientoF.Recordset
    If .RecordCount <= 30 Then
      If TxtDetalle.Text <> Ninguno Then Producto = TxtDetalle.Text
      TxtDetalle.Visible = False
    ' MsgBox CCur(TextCant.Text) * CCur(TextVUnit.Text)
      Real1 = Redondear(CCur(TextCant.Text) * CCur(TextVUnit.Text), 2)
      'Real2 = Redondear(Real1 * Val(TextDesc1.Text) / 100, 2)
      If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
      If TipoFactura = "NV" Then Real3 = 0
      LabelVTotal.Caption = Format$(Real1, "#,##0.00")
      SetAdoAddNew "Asiento_F"
      SetAdoFields "CODIGO", Codigos
      SetAdoFields "CODIGO_L", CodigoL
      SetAdoFields "PRODUCTO", Producto
      SetAdoFields "REP", 0
      SetAdoFields "TONELAJE", Val(TxtTonelaje.Text)
      SetAdoFields "CANT", CCur(TextCant.Text)
      SetAdoFields "PRECIO", CCur(TextVUnit.Text)
      SetAdoFields "TOTAL", Real1
      SetAdoFields "Total_Desc", Real2
      SetAdoFields "Total_IVA", Real3
      SetAdoFields "Cta", Cta_Ventas
      SetAdoFields "Item", NumEmpresa
      SetAdoFields "CodigoU", CodigoUsuario
      'If Val(TextComEjec.Text) > 0 Then
      '   SetAdoFields "Cod_Ejec", CodigoEjec
      '   SetAdoFields "Porc_C", Redondear(Val(TextComEjec.Text) / 100, 4)
      'End If
      SetAdoUpdate
      Calculos_Totales_Factura FA
      TextVUnit.Text = ""
      DCArticulo.SetFocus
   Else
      MsgBox "Ya no se puede ingresar más datos."
      Command1.SetFocus
   End If
   End With
   End If
End Sub

Private Sub TxtDAU_GotFocus()
  MarcarTexto TxtDAU
End Sub

Private Sub TxtDAU_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDeclaracion_GotFocus()
  MarcarTexto TxtDeclaracion
End Sub

Private Sub TxtDeclaracion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTextoFinal TxtDetalle
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then TxtDetalle.Visible = False
End Sub

Private Sub TxtDetalle_LostFocus()
  TxtDetalle.Visible = False
  TextCant.SetFocus
End Sub

Private Sub TxtFactCom_GotFocus()
   MarcarTexto TxtFactCom
End Sub

Private Sub TxtFactCom_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFUE_GotFocus()
   MarcarTexto TxtFUE
End Sub

Private Sub TxtFUE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtGuiaRem_GotFocus()
  MarcarTexto TxtGuiaRem
End Sub

Private Sub TxtGuiaRem_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTonelaje_GotFocus()
  MarcarTexto TxtTonelaje
End Sub

Private Sub TxtTonelaje_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTonelaje_LostFocus()
  TextoValido TxtTonelaje
End Sub
