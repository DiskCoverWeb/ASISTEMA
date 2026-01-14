VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form ListFact 
   BackColor       =   &H80000002&
   Caption         =   "FACTURACION: Listar Factura"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11280
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmEjecutivo 
      BackColor       =   &H00404080&
      Caption         =   "| SELECCIONE EL EJECUTIVO DE VENTA |"
      ForeColor       =   &H0000FFFF&
      Height          =   3690
      Left            =   14595
      TabIndex        =   57
      Top             =   3780
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "&Aceptar"
         Height          =   855
         Left            =   6510
         Picture         =   "ListFact.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1575
         Width           =   1800
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000080FF&
         Caption         =   "&Salir"
         Height          =   855
         Left            =   6510
         Picture         =   "ListFact.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2520
         Width           =   1800
      End
      Begin VB.OptionButton OpcAbonos 
         BackColor       =   &H00004080&
         Caption         =   "Abonos"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6510
         TabIndex        =   61
         Top             =   1155
         Width           =   1170
      End
      Begin VB.OptionButton OpcFactura 
         BackColor       =   &H00004080&
         Caption         =   "Factura"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6510
         TabIndex        =   60
         Top             =   735
         Width           =   1170
      End
      Begin VB.OptionButton OpcAmbos 
         BackColor       =   &H00004080&
         Caption         =   "Factura y Abonos"
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6510
         TabIndex        =   59
         Top             =   315
         Value           =   -1  'True
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo DCEjecutivo 
         Bindings        =   "ListFact.frx":0D0C
         DataSource      =   "AdoEjecutivo"
         Height          =   3105
         Left            =   210
         TabIndex        =   58
         Top             =   315
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   5477
         _Version        =   393216
         Style           =   1
         BackColor       =   12640511
         Text            =   ""
      End
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "ListFact.frx":0D27
      DataSource      =   "AdoArticulo"
      Height          =   3540
      Left            =   525
      TabIndex        =   66
      Top             =   5880
      Visible         =   0   'False
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   6138
      _Version        =   393216
      Style           =   1
      BackColor       =   12648384
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqClaveAcceso 
      Alignment       =   1  'Right Justify
      Caption         =   "Clave de Accceso:"
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   1470
      Width           =   2220
   End
   Begin VB.TextBox TxtDetalle 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3270
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   67
      Text            =   "ListFact.frx":0D41
      Top             =   5670
      Visible         =   0   'False
      Width           =   13770
   End
   Begin VB.ListBox LstBox 
      Height          =   1620
      Left            =   14490
      TabIndex        =   75
      Top             =   840
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "ListFact.frx":0D47
      Height          =   2010
      Left            =   105
      TabIndex        =   65
      ToolTipText     =   "<Alt+F9> Cambia el Producto, <Alt+F10> Cambia la Bodega, <Ctrl+S> Actualiza la Serie, "
      Top             =   5460
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   3545
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648384
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin VB.Frame FrmTotalAsiento 
      Height          =   540
      Left            =   105
      TabIndex        =   69
      Top             =   10500
      Visible         =   0   'False
      Width           =   7890
      Begin VB.Label LabelHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   5985
         TabIndex        =   74
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label LabelDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   4095
         TabIndex        =   73
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label LblDiferencia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   330
         Left            =   1155
         TabIndex        =   72
         Top             =   105
         Width           =   1800
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diferencia "
         Height          =   330
         Left            =   105
         TabIndex        =   71
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTALES "
         Height          =   330
         Left            =   3045
         TabIndex        =   70
         Top             =   105
         Width           =   1065
      End
   End
   Begin VB.TextBox TxtXML 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   9765
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   64
      Top             =   2310
      Width           =   12375
   End
   Begin VB.Frame FrmTotales 
      Height          =   750
      Left            =   105
      TabIndex        =   39
      Top             =   9660
      Width           =   12930
      Begin VB.Label LabelSaldoAct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   11340
         TabIndex        =   50
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo Actual"
         Height          =   330
         Left            =   11340
         TabIndex        =   51
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label LabelTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   9765
         TabIndex        =   52
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Facturado"
         Height          =   330
         Left            =   9765
         TabIndex        =   53
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label LabelServicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   8085
         TabIndex        =   48
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Subtotal Servicio"
         Height          =   330
         Left            =   8085
         TabIndex        =   49
         Top             =   105
         Width           =   1695
      End
      Begin VB.Label LabelIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6615
         TabIndex        =   40
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " I.V.A."
         Height          =   330
         Left            =   6615
         TabIndex        =   41
         Top             =   105
         Width           =   1485
      End
      Begin VB.Label LabelSubTotalFA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4935
         TabIndex        =   56
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SubTotal"
         Height          =   330
         Left            =   4935
         TabIndex        =   55
         Top             =   105
         Width           =   1695
      End
      Begin VB.Label LabelDesc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3360
         TabIndex        =   46
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento"
         Height          =   330
         Left            =   3360
         TabIndex        =   47
         Top             =   105
         Width           =   1590
      End
      Begin VB.Label LabelConIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1680
         TabIndex        =   42
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label333 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Subtotal con IVA"
         Height          =   330
         Left            =   1680
         TabIndex        =   43
         Top             =   105
         Width           =   1695
      End
      Begin VB.Label LabelSubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   0
         TabIndex        =   44
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Subtotal sin IVA"
         Height          =   330
         Left            =   0
         TabIndex        =   45
         Top             =   105
         Width           =   1695
      End
   End
   Begin VB.TextBox TxtObs 
      Height          =   330
      Left            =   1470
      MaxLength       =   100
      TabIndex        =   23
      ToolTipText     =   "<Ctrl+O> Coloca una Observacion a la factura"
      Top             =   3780
      Width           =   13035
   End
   Begin VB.TextBox TxtAutorizacion 
      Height          =   330
      Left            =   9030
      MaxLength       =   50
      TabIndex        =   36
      ToolTipText     =   "<Ctrl-A> Grabar Autorizacion Electronica manualmente"
      Top             =   1470
      Width           =   5370
   End
   Begin VB.TextBox TxtClaveAcceso 
      Height          =   330
      Left            =   2310
      MaxLength       =   50
      TabIndex        =   35
      ToolTipText     =   "<Ctrl+S> Volver a Procesar el Documento, <Ctrl+R> Recalcular Saldo Factura"
      Top             =   1470
      Width           =   5370
   End
   Begin VB.Frame Frame1 
      Caption         =   "En Bloque"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1065
      Left            =   14490
      TabIndex        =   20
      Top             =   735
      Width           =   5790
      Begin VB.CheckBox CheqSinCodigo 
         Caption         =   "Imprimir sin Codigo de Alumna"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2415
         TabIndex        =   38
         Top             =   630
         Width           =   1590
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Actualizar Alumnos"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         TabIndex        =   34
         ToolTipText     =   "Cambio de fecha de vencimiento en un rango de facturas"
         Top             =   630
         Width           =   1590
      End
      Begin VB.CheckBox CheqSoloCopia 
         Caption         =   "Imprimir Solo Copia"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   30
         Top             =   630
         Width           =   1170
      End
      Begin VB.CheckBox CheqMatricula 
         Caption         =   "Sin Deuda Pendiente"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   29
         Top             =   630
         Width           =   960
      End
      Begin VB.OptionButton OpcDes 
         Height          =   345
         Left            =   3990
         Picture         =   "ListFact.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   210
         Width           =   330
      End
      Begin VB.OptionButton OpcAsc 
         Height          =   345
         Left            =   3675
         Picture         =   "ListFact.frx":1286
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox TextFHasta 
         Height          =   330
         Left            =   2625
         TabIndex        =   10
         Text            =   "0"
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox TextFDesde 
         Height          =   330
         Left            =   840
         TabIndex        =   8
         Text            =   "0"
         Top             =   210
         Width           =   960
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   4410
         TabIndex        =   33
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
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
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta:"
         Height          =   330
         Left            =   1890
         TabIndex        =   9
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde:"
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   14295
      Begin MSDataListLib.DataCombo DCFact 
         Bindings        =   "ListFact.frx":17AC
         DataSource      =   "AdoFactList"
         Height          =   360
         Left            =   6195
         TabIndex        =   6
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "000000000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCTipo 
         Bindings        =   "ListFact.frx":17C6
         DataSource      =   "AdoTipo"
         Height          =   360
         Left            =   1785
         TabIndex        =   2
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "FA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCSerie 
         Bindings        =   "ListFact.frx":17DC
         DataSource      =   "AdoSerie"
         Height          =   360
         Left            =   3465
         TabIndex        =   4
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "001001"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAutorizacion 
         Height          =   330
         Left            =   105
         Top             =   735
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
      Begin MSAdodcLib.Adodc AdoSerie 
         Height          =   330
         Left            =   2310
         Top             =   735
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSAdodcLib.Adodc AdoFactList 
         Height          =   330
         Left            =   5880
         Top             =   735
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
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
         Caption         =   "FactList"
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
      Begin MSAdodcLib.Adodc AdoTipo 
         Height          =   330
         Left            =   3990
         Top             =   735
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Tipo"
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
      Begin VB.Label LabelCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9999999999"
         Height          =   330
         Left            =   10815
         TabIndex        =   32
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9999"
         Height          =   330
         Left            =   9240
         TabIndex        =   54
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label LabelFechaPe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31/12/2009"
         Height          =   330
         Left            =   7980
         TabIndex        =   31
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label LabelEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PENDIENTE"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   12390
         TabIndex        =   28
         Top             =   210
         Width           =   1800
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Secuencial No."
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4725
         TabIndex        =   5
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Documento"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Serie"
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2730
         TabIndex        =   3
         Top             =   210
         Width           =   750
      End
   End
   Begin TabDlg.SSTab SSTabDetalle 
      Height          =   390
      Left            =   105
      TabIndex        =   27
      Top             =   5040
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   688
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DETALLE DE FACTURA"
      TabPicture(0)   =   "ListFact.frx":17F3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "ABONOS DE LA FACTURA"
      TabPicture(1)   =   "ListFact.frx":180F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "GUIA DE REMISION"
      TabPicture(2)   =   "ListFact.frx":182B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "CONTABILIZACION"
      TabPicture(3)   =   "ListFact.frx":1847
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&S"
      Height          =   330
      Left            =   105
      TabIndex        =   24
      Top             =   9660
      Width           =   435
   End
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   2415
      Top             =   6615
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoDiarioCaja 
      Height          =   330
      Left            =   2415
      Top             =   7245
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   210
      Top             =   6615
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   210
      Top             =   6930
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Left            =   210
      Top             =   7245
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13560
      Top             =   9675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":1863
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":213D
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":2A17
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":51C9
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":5A5B
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":5D75
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":664F
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":6969
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A1BEB
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A1F05
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A28B3
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A318D
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A34A7
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A3D81
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A465B
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A4D01
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A531C
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A5BF6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A70BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":A7D15
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":142F97
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":1439BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":143E0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":149147
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":14FC54
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":1566F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":15D661
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":15D97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":15DC95
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":15E8E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListFact.frx":15F539
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoEjecutivo 
      Height          =   330
      Left            =   2415
      Top             =   6930
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSComctlLib.Toolbar TBarFactura 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir de listar facturas"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir_Factura"
            Object.ToolTipText     =   "Imprimir Factura Individual"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_Individual"
                  Text            =   "Factura Individual"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_En_Bloque"
                  Text            =   "Facturas en Bloque"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_PV"
                  Text            =   "Imprime en impresora P.V."
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_NC"
                  Text            =   "Notas de Crédito"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_Recibos"
                  Text            =   "Recibos en Bloque"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRN_Guia_R"
                  Text            =   "Guia de Remision"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Enviar_Mail"
            Object.ToolTipText     =   "Reenviar: Factura,NC o Guia por Email"
            ImageIndex      =   25
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mail_FA"
                  Text            =   "Enviar Email Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mail_NC"
                  Text            =   "Enviar Email Nota de Credito"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mail_GR"
                  Text            =   "Enviar Email Guia de Remision"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bajar_XML"
            Object.ToolTipText     =   "Descargar: Factura,NC o Guia por Email"
            ImageIndex      =   30
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "XML_FA"
                  Text            =   "Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "XML_LC"
                  Text            =   "Liquidacion de Compras"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "XML_NC"
                  Text            =   "Nota de Credito"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "XML_GR"
                  Text            =   "Guia de Remision"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Generar_PDF"
            Object.ToolTipText     =   "Genera el PDF de: Factura, Nota de Credito o Guia de Remision"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_FA"
                  Text            =   "Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_LC"
                  Text            =   "Liquidacion de Compras"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_NC"
                  Text            =   "Nota de Credito"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_GR"
                  Text            =   "Guia de Remision"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_OP"
                  Text            =   "Orden de Produccion"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PDF_DO"
                  Text            =   "Donacion"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cambio_Emision_Facturas"
            Object.ToolTipText     =   "Cambia la Fecha en un rango de Facturas"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cambio_Vencimiento_Facturas"
            Object.ToolTipText     =   "Cambia Fecha de Vencimiento de Facturas"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cambia_Autorizacion_Facturas"
            Object.ToolTipText     =   "Cambio de Autorizacion en un rango de facturas"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cambia_Numero_de_Facturas"
            Object.ToolTipText     =   "Cambia el numero de Facturas de un rango de Facturas"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reprocesar_Saldos_Facturas"
            Object.ToolTipText     =   "Reprocesar Saldos de Facturas"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar_Facturas"
            Object.ToolTipText     =   "Elimina Facturas en un rango determinado"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Revertir_Facturas"
            Object.ToolTipText     =   "Reversar Facturas procesadas"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Actualizar_Representantes"
            Object.ToolTipText     =   "Actualiza Representantes en un rango de Facturas"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anular_Factura"
            Object.ToolTipText     =   "Anular Factura"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anular_en_masa"
            Object.ToolTipText     =   "Anular Facturas en masa"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Nota_Credito"
            Object.ToolTipText     =   "Nota de Crédito"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Separador_SRI"
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Liberar_FA_SRI"
            Object.ToolTipText     =   "Libera rango de Facturas SRI"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Volver_Autorizar_SRI"
            Object.ToolTipText     =   "Autorizar documentos pendiente"
            ImageIndex      =   18
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SRIFactAct"
                  Text            =   "Factura Actual"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SRIAutNC"
                  Text            =   "Nota de Credito"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SRIFactPend"
                  Text            =   "Facturas Pendientes"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GuiaR"
                  Text            =   "Guia de Remision"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ejecutivo"
            Object.ToolTipText     =   "Actualiza Ejecutivo de Venta"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Kardex"
            Object.ToolTipText     =   "Actualiza Salida en Kardex"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel el detalle"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ayuda"
            Object.ToolTipText     =   "Ayuda de Comandos Automaticos"
            ImageIndex      =   31
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   13860
         TabIndex        =   78
         Top             =   0
         Width           =   2430
         Begin VB.CheckBox CheqAutxRangos 
            Caption         =   "Autorizar por Rangos"
            Height          =   225
            Left            =   105
            TabIndex        =   79
            Top             =   210
            Width           =   2220
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   10185
         TabIndex        =   77
         Top             =   525
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1111
         ButtonWidth     =   609
         ButtonHeight    =   1005
         Appearance      =   1
         _Version        =   393216
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
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
      ForeColor       =   &H00C0FFFF&
      Height          =   1380
      Left            =   105
      TabIndex        =   13
      Top             =   2310
      Width           =   9570
   End
   Begin VB.Label LabelBultos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   330
      Left            =   11235
      TabIndex        =   19
      Top             =   4620
      Width           =   960
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. de Bultos:"
      Height          =   330
      Left            =   9765
      TabIndex        =   18
      Top             =   4620
      Width           =   1485
   End
   Begin VB.Label LabelTransp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1470
      TabIndex        =   16
      Top             =   4200
      Width           =   13035
   End
   Begin VB.Label LabelVendedor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   10290
      TabIndex        =   14
      Top             =   1890
      Width           =   9990
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorizacion"
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7770
      TabIndex        =   37
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nota:"
      Height          =   330
      Left            =   105
      TabIndex        =   15
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Observación:"
      Height          =   330
      Left            =   105
      TabIndex        =   17
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label LabelCliente 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   330
      Left            =   945
      TabIndex        =   12
      Top             =   1890
      Width           =   9255
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1470
      TabIndex        =   21
      Top             =   4620
      Width           =   8205
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Entregado en:"
      Height          =   330
      Left            =   105
      TabIndex        =   22
      Top             =   4620
      Width           =   1380
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cliente:"
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   1890
      Width           =   855
   End
End
Attribute VB_Name = "ListFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'&H8000000F&
'&H80000005&
'Variables Globales

Public Sub Listar_Factura_NotaVentas()
  FA.TC = DCTipo
  FA.Serie = DCSerie
  sSQL = "SELECT Factura, Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND Serie = '" & FA.Serie & "' "
  If Modulo = "EJECUTIVOS" Then sSQL = sSQL & "AND Cod_Ejec = '" & CodigoUsuario & "' "
  sSQL = sSQL & "ORDER BY Factura "
  SelectDB_Combo DCFact, AdoFactList, sSQL, "Factura"
  If AdoFactList.Recordset.RecordCount > 0 Then
     AdoFactList.Recordset.MoveLast
     DCFact = AdoFactList.Recordset.fields("Factura")
     FA.Factura = AdoFactList.Recordset.fields("Factura")
     FA.Autorizacion = AdoFactList.Recordset.fields("Autorizacion")
     TextFDesde = DCFact
     TextFHasta = DCFact
  Else
     DCFact = "0"
     TextFDesde = DCFact
     TextFHasta = DCFact
     MsgBox "Esta Empresa no ha comenzado a Facturar," _
          & "en este Tipo de Facturacion (" & FA.TC & "), " _
          & "ingrese su primer Documento."
     Unload Me
  End If
End Sub

Public Sub Revertir_Facturas()
  DGDetalle.Visible = False
  If Val(TextFDesde) <= 0 Then TextFDesde = DCFact
  If Val(TextFHasta) <= 0 Then TextFHasta = DCFact
  TextoValido TextFDesde
  TextoValido TextFHasta
  Autorizacion = TxtAutorizacion
  SerieFactura = DCSerie
  TipoDoc = DCTipo
  Factura_Desde = Val(TextFDesde)
  Factura_Hasta = Val(TextFHasta)
  If Factura_Desde > Factura_Hasta Then Factura_Hasta = Factura_Desde
  Mensajes = "Esta Seguro de Reversar desde " & vbCrLf _
           & "La Factura No. " & TextFDesde & " hasta la " & TextFHasta & vbCrLf _
           & "en bloque "
  Titulo = "Formulario de Eliminación"
  If BoxMensaje = vbYes Then
  RatonReloj
  Contador = 0
  sSQL = "SELECT DF.*,C.Grupo " _
       & "FROM Detalle_Factura As DF,Clientes As C " _
       & "WHERE DF.Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND DF.Autorizacion = '" & Autorizacion & "' " _
       & "AND DF.Serie = '" & SerieFactura & "' " _
       & "AND DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.TC = '" & TipoDoc & "' " _
       & "AND DF.CodigoC = C.Codigo " _
       & "ORDER BY DF.CodigoC,DF.Codigo "
  Select_Adodc AdoDetalle, sSQL
  With AdoDetalle.Recordset
   If .RecordCount > 0 Then
       MsgBox "Se va a revertir: " & .RecordCount & " item de facturacion."
       Do While Not .EOF
          Contador = Contador + 1
          ListFact.Caption = Format(Contador / .RecordCount, "00%")
          I = .fields("Mes_No")
          Codigo2 = .fields("Ticket")
          Codigo3 = .fields("Mes")
          
          If I <= 0 Then I = Month(.fields("Fecha"))
          If Len(Codigo2) < 4 Then Codigo1 = CStr(Year(.fields("Fecha")))
          If Len(Codigo3) < 3 Then Codigo3 = MesesLetras(CInt(I))
          
          sSQL = "DELETE * " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Periodo = '" & Codigo2 & "' " _
               & "AND Codigo_Inv = '" & .fields("Codigo") & "' " _
               & "AND Codigo = '" & .fields("CodigoC") & "' " _
               & "AND Num_Mes = " & I & " " _
               & "AND Item = '" & NumEmpresa & "' "
          Ejecutar_SQL_SP sSQL
          
          SetAdoAddNew "Clientes_Facturacion"
          SetAdoFields "T", Normal
          SetAdoFields "Codigo", .fields("CodigoC")
          SetAdoFields "Codigo_Inv", .fields("Codigo")
          SetAdoFields "Valor", .fields("Precio")
          SetAdoFields "Descuento", .fields("Total_Desc")
          SetAdoFields "Descuento2", .fields("Total_Desc2")
          SetAdoFields "GrupoNo", .fields("Grupo")
          SetAdoFields "Num_Mes", I
          SetAdoFields "Mes", Codigo3
          SetAdoFields "Fecha", .fields("Fecha")
          SetAdoFields "Periodo", Codigo2
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "CodigoU", CodigoUsuario
          SetAdoUpdate
         'MsgBox sSQL
         .MoveNext
       Loop
   End If
  End With
  sSQL = "DELETE * " _
       & "FROM Facturas " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Detalle_Factura " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Trans_Abonos " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = '" & TipoDoc & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If DCTipo.Text = "NV" Then
     sSQL = sSQL & "AND TC = 'NV' "
  Else
     sSQL = sSQL & "AND TC NOT IN ('C','P') "
  End If
  sSQL = sSQL & "ORDER BY Factura DESC "
  Select_Adodc AdoDiarioCaja, sSQL
  With AdoDiarioCaja.Recordset
   If .RecordCount > 0 Then
       Factura_Hasta = .fields("Factura") + 1
       If DCTipo.Text = "NV" Then
          WriteSetDataNum "Nota Ventas", False, Factura_Hasta
       Else
          WriteSetDataNum "Facturas", False, Factura_Hasta
       End If
   End If
  End With
  Listar_Factura_NotaVentas
  DGDetalle.Visible = True
  RatonNormal
  MsgBox "Proceso Terminado"
  Unload ListFact
  End If
End Sub

Public Sub Imprimir_Recibos(Optional CheqSinCodigo As Boolean)
  TextoValido TxtObs, , True
  TextoValido TextFDesde
  TextoValido TextFHasta
  
  If DCFact = "" Then DCFact = "0"
  If Val(TextFDesde) <= 0 Then TextFDesde = DCFact
  If Val(TextFHasta) <= 0 Then TextFHasta = DCFact
  FA.Tipo_PRN = "FM"
  FA.Desde = Val(TextFDesde)
  FA.Hasta = Val(TextFHasta)
  FA.TC = DCTipo
  FA.Serie = DCSerie
  If TxtObs = "" Then TxtObs = Ninguno
  SQL2 = "UPDATE Facturas " _
       & "SET Observacion = '" & TxtObs & "' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = '" & FA.TC & "' " _
       & "AND Serie = '" & FA.Serie & "' " _
       & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
  Ejecutar_SQL_SP SQL2
  
  If (FA.Hasta - FA.Desde) >= 0 Then
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 5, Pos_Yo = 0.5 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 0 "
     Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 5, Pos_Yo = 1.5 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 1 "
     Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 2, Pos_Yo = 0.5 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 50 "
     Ejecutar_SQL_SP sSQL
     SQL2 = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " " _
          & "ORDER BY Factura,Cod_CxC "
     Select_Adodc AdoFactura, SQL2
     If AdoFactura.Recordset.RecordCount > 0 Then
        FA.Cod_CxC = AdoFactura.Recordset.fields("Cod_CxC")
        Lineas_De_CxC FA
        FA.TC = DCTipo
        FA.Serie = DCSerie
        Bandera = False
        Evaluar = True
       'MsgBox FA.TC & "........"
        If FA.Desde <= FA.Hasta Then
           If CheqSoloCopia.value = 1 Then
              Imprimir_Facturas_Copias_CxC ListFact, AdoFactura, AdoDetalle, Factura_Desde, Factura_Hasta, FA, True, CBool(CheqMatricula.value), CBool(OpcAsc.value)
           Else
              Imprimir_Facturas_CxC ListFact, FA, True, CBool(CheqMatricula.value), True, CBool(OpcAsc.value), , CBool(CheqSinCodigo)
           End If
        End If
        Facturas_Impresas FA
     End If
  Else
    MsgBox "No se puede imprimir el rando de Facturas"
  End If
End Sub

Public Sub Imprimir_NC()
  If DCFact.Text = "" Then DCFact.Text = "0"
  If Val(TextFDesde.Text) <= 0 Then TextFDesde.Text = DCFact.Text
  If Val(TextFHasta.Text) <= 0 Then TextFHasta.Text = DCFact.Text
  TextoValido TextFDesde
  TextoValido TextFHasta
  Factura_Desde = Val(TextFDesde.Text)
  Factura_Hasta = Val(TextFHasta.Text)
  FA.TC = DCTipo
  FA.Serie = DCSerie
  FA.Autorizacion = TxtAutorizacion
  Control_Procesos "I", "Reimpresion de Facturas desde la " & Factura_Desde & " a la " & Factura_Hasta
  If (Factura_Hasta - Factura_Desde) >= 0 Then
     SQL2 = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " " _
          & "ORDER BY Cod_CxC "
     Select_Adodc AdoFactura, SQL2
     If AdoFactura.Recordset.RecordCount > 0 Then
        FA.Cod_CxC = AdoFactura.Recordset.fields("Cod_CxC")
       'Lineas de CxC Clientes
       'Lineas_De_CxC FA
'''        FA.TC = DCTipo
'''        FA.Serie = DCSerie
'''        FA.Autorizacion = TxtAutorizacion
        Bandera = False
        Evaluar = True
        Imprimir_Nota_Credito ListFact, AdoFactura, AdoDetalle, Factura_Desde, Factura_Hasta, FA
     End If
  Else
    MsgBox "No se puede imprimir el rando de Notas de Crédito"
  End If
End Sub

Public Sub Cambia_Vencimiento_Facturas()
  If ClaveAdministrador Then
     FechaValida MBFecha
     Mifecha = BuscarFecha(MBFecha)
     MiMes = MesesLetras(Month(MBFecha))
     If DCFact.Text = "" Then DCFact.Text = "0"
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     Control_Procesos "F", "Cambio de Fecha de vencimiento desde " & Factura_Desde & " a la " & Factura_Hasta, 0
  If (Factura_Hasta - Factura_Desde) >= 0 Then
     RatonReloj
     SQL2 = "UPDATE Facturas " _
          & "SET Fecha_V = #" & Mifecha & "# " _
          & "WHERE Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " " _
          & "AND TC = '" & DCTipo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Ejecutar_SQL_SP SQL2
     RatonNormal
  Else
    MsgBox "No se puede procesar el rando de Facturas"
  End If
  End If
End Sub

Public Sub Eliminar_Facturas()
  If Val(TextFDesde) <= 0 Then TextFDesde = "0"
  If Val(TextFHasta) <= 0 Then TextFHasta = "0"
  Autorizacion = TrimStrg(TxtAutorizacion)
  SerieFactura = TrimStrg(DCSerie)
  TextoValido TextFDesde
  TextoValido TextFHasta
  Factura_Desde = Val(TextFDesde.Text)
  Factura_Hasta = Val(TextFHasta.Text)
  If Factura_Desde > Factura_Hasta Then Factura_Hasta = Factura_Desde
  Mensajes = "Esta Seguro de Eliminar desde " & vbCrLf _
           & "La Factura No. " & TextFDesde.Text & " hasta la " & TextFHasta.Text & vbCrLf _
           & "en bloque "
  Titulo = "Formulario de Eliminación"
  If BoxMensaje = vbYes Then
  RatonReloj
  TipoDoc = "AND TC = '" & DCTipo & "' "
  TipoComp = "AND TP = '" & DCTipo & "' "
  sSQL = "DELETE * " _
       & "FROM Facturas " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & TipoDoc
  'MsgBox sSQL
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Detalle_Factura " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & TipoDoc
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Trans_Abonos " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & TipoComp
  Ejecutar_SQL_SP sSQL
  sSQL = "DELETE * " _
       & "FROM Trans_Kardex " _
       & "WHERE Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & TipoDoc
  Ejecutar_SQL_SP sSQL
  
 'MsgBox sSQL
  sSQL = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If DCTipo.Text = "NV" Then
     sSQL = sSQL & "AND TC = 'NV' "
  Else
     sSQL = sSQL & "AND TC NOT IN ('C','P') "
  End If
  sSQL = sSQL & "ORDER BY Factura DESC "
  Select_Adodc AdoDiarioCaja, sSQL
  With AdoDiarioCaja.Recordset
   If .RecordCount > 0 Then
       Factura_Hasta = .fields("Factura") + 1
       If DCTipo.Text = "NV" Then
          WriteSetDataNum "Nota Ventas", False, Factura_Hasta
       Else
          WriteSetDataNum "Facturas", False, Factura_Hasta
       End If
   End If
  End With
  Listar_Factura_NotaVentas
  RatonNormal
  End If
End Sub

Private Sub Command1_Click()
   With AdoEjecutivo.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente Like '" & DCEjecutivo & "' ")
        If Not .EOF Then
           FA.Cod_Ejec = .fields("Codigo")
        Else
           MsgBox "Ejecutivo de Venta no asignado"
           FA.Cod_Ejec = Ninguno
        End If
    Else
        MsgBox "No hay Ejecutivos de Venta asignados"
        FA.Cod_Ejec = Ninguno
    End If
   End With
   
   If OpcFactura.value Then
      sSQL = "UPDATE Facturas " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
      
      sSQL = "UPDATE Detalle_Factura " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
   ElseIf OpcAbonos.value Then
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
   Else
      sSQL = "UPDATE Facturas " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
      
      sSQL = "UPDATE Detalle_Factura " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
     
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Cod_Ejec = '" & FA.Cod_Ejec & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Factura = '" & FA.Factura & "' "
      Ejecutar_SQL_SP sSQL
   End If
   FrmEjecutivo.Visible = False
   MsgBox "Proceso realizado con exito, liste nuevamente la factura"
End Sub

Private Sub Command2_Click()
   FrmEjecutivo.Visible = False
End Sub

Private Sub Command7_Click()
 'MsgBox sSQL
  If ClaveAdministrador Then
     RatonReloj
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     Contador = 0
     If (Factura_Hasta - Factura_Desde) >= 0 Then
        sSQL = "SELECT F.Factura,F.Serie,F.TC,C.Codigo,C.CI_RUC,CM.Lugar_Trabajo_R,CM.Representante,CM.Cedula_R,CM.Telefono_R " _
             & "FROM Facturas As F, Clientes As C, Clientes_Matriculas As CM " _
             & "WHERE F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.T <> '" & Anulado & "' " _
             & "AND F.TC = '" & FA.TC & "' " _
             & "AND F.Serie = '" & FA.Serie & "' " _
             & "AND F.Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND F.Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " " _
             & "AND F.Item = CM.Item " _
             & "AND F.Periodo = CM.Periodo " _
             & "AND F.CodigoC = C.Codigo " _
             & "AND F.CodigoC = CM.Codigo " _
             & "ORDER BY F.Factura "
        Select_Adodc AdoDiarioCaja, sSQL
        RatonReloj
        With AdoDiarioCaja.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                ListFact.Caption = Format$(Contador / .RecordCount, "00.00%") & ", Codigo: " & .fields("CI_RUC")
                Estudiante_DBF.cedular = .fields("Cedula_R")
                Estudiante_DBF.fonopaga = .fields("Telefono_R")
                Estudiante_DBF.pagador = .fields("Representante")
                Estudiante_DBF.direcpaga = .fields("Lugar_Trabajo_R")
                Estudiante_DBF.codest = .fields("CI_RUC")
                Estudiante_DBF.cedula = .fields("CI_RUC")
                'Actualizar_Pagos
                Contador = Contador + 1
               .MoveNext
             Loop
         End If
        End With
     End If
     RatonNormal
     MsgBox "Proceso Terminado"
     DCFact.SetFocus
  End If
End Sub

Public Sub Cambia_Numero_de_Facturas()
Dim FormCaption As String
    FormCaption = ListFact.Caption
    If Val(TextFDesde) <= 0 Then TextFDesde = "0"
    If Val(TextFHasta) <= 0 Then TextFHasta = "0"
    TextoValido TextFDesde
    TextoValido TextFHasta
    Factura_Desde = Val(TextFDesde)
    Factura_Hasta = Val(TextFHasta)
    Control_Procesos "F", "Anulacion de Facturas en Lotes, Desde: " & Factura_Desde & " a la " & Factura_Hasta
    Numero = Val(InputBox("Ingrese el Numero de la" & vbCrLf & "Factura a Cambiar:", "CAMBIO DE NUMERO DE FACTURAS", "0"))
    
    Mensajes = "Cambiar el numero de la(s) Factura(s)" & vbCrLf _
             & "Desde: " & Format$(Factura_Desde, "0000000") & ", hasta: " & Format$(Factura_Hasta, "0000000") & vbCrLf _
             & "Por el Numero desde: " & Format$(Numero, "0000000") & vbCrLf _
             & "Autorizacion: " & FA.TC & "-" & FA.Serie & " No. " & FA.Autorizacion
    Titulo = "Formulario de Grabación."
    If BoxMensaje = vbYes And Numero > 0 And Factura_Desde > 0 And Factura_Hasta > 0 Then
       For Factura_No = Factura_Desde To Factura_Hasta
           FA.CodigoC = ""
           sSQL = "SELECT CodigoC " _
                & "FROM Facturas " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' "
           Select_Adodc AdoDiarioCaja, sSQL
           If AdoDiarioCaja.Recordset.RecordCount > 0 Then
              FA.CodigoC = AdoDiarioCaja.Recordset.fields("CodigoC")
           End If
           If FA.CodigoC <> "" Then
              sSQL = "UPDATE Facturas " _
                   & "SET Factura = " & Numero & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND TC = '" & FA.TC & "' " _
                   & "AND Serie = '" & FA.Serie & "' " _
                   & "AND Autorizacion = '" & FA.Autorizacion & "' "
              Ejecutar_SQL_SP sSQL
              
              sSQL = "UPDATE Detalle_Factura " _
                   & "SET Factura = " & Numero & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND TC = '" & FA.TC & "' " _
                   & "AND Serie = '" & FA.Serie & "' " _
                   & "AND Autorizacion = '" & FA.Autorizacion & "' "
              Ejecutar_SQL_SP sSQL
             
              sSQL = "UPDATE Trans_Abonos " _
                   & "SET Factura = " & Numero & " " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Factura = " & Factura_No & " " _
                   & "AND TP = '" & FA.TC & "' " _
                   & "AND Serie = '" & FA.Serie & "' " _
                   & "AND Autorizacion = '" & FA.Autorizacion & "' "
              Ejecutar_SQL_SP sSQL
             'Anulamos la misma factura anterior
              SetAdoAddNew "Facturas"
              SetAdoFields "T", Anulado
              SetAdoFields "TC", FA.TC
              SetAdoFields "Factura", Factura_No
              SetAdoFields "CodigoC", FA.CodigoC
              SetAdoFields "Cod_CxC", FA.Cod_CxC
              SetAdoFields "Cta_CxP", FA.Cta_CxP
              SetAdoFields "Serie", FA.Serie
              SetAdoFields "Autorizacion", FA.Autorizacion
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "Periodo", Periodo_Contable
              SetAdoUpdate
           End If
           ListFact.Caption = "Anulando la Factura No. " & Factura_No & " y cambiandola por " & Numero
           Numero = Numero + 1
       Next Factura_No
       ListFact.Caption = FormCaption
       MsgBox "Proceso Terminado"
    End If
End Sub

Public Sub Volver_Autorizar(Optional ConMensaje As Boolean)
Dim ActualizarCliente As Boolean
    ''       MsgBox MensajeNoAutorizarCE
    TxtXML.Text = ""
    If CFechaLong(LabelFechaPe) <= CFechaLong(Fecha_CE) Then
       If FA.T <> "A" Then
          If ConMensaje Then
             Mensajes = "Liberar Documento: " & FA.TC & "-" & FA.Serie & " No. " & FA.Factura & " y volver a autorizar"
             Titulo = "Formulario de Actualizacion."
             If BoxMensaje = vbYes Then
                SQL2 = "UPDATE Facturas " _
                     & "SET Autorizacion = '" & RUC & "', Clave_Acceso = '" & Ninguno & "', Estado_SRI = '" & Ninguno & "' " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND TC = '" & FA.TC & "' " _
                     & "AND Serie = '" & FA.Serie & "' " _
                     & "AND Factura = " & FA.Factura & " "
                Ejecutar_SQL_SP SQL2
                
                SQL2 = "UPDATE Detalle_Factura " _
                     & "SET Autorizacion = '" & RUC & "' " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND TC = '" & FA.TC & "' " _
                     & "AND Serie = '" & FA.Serie & "' " _
                     & "AND Factura = " & FA.Factura & " "
                Ejecutar_SQL_SP SQL2
                 
                SQL2 = "UPDATE Trans_Abonos " _
                     & "SET Autorizacion = '" & RUC & "' " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND TP = '" & FA.TC & "' " _
                     & "AND Serie = '" & FA.Serie & "' " _
                     & "AND Factura = " & FA.Factura & " "
                Ejecutar_SQL_SP SQL2
                 
                FA.Autorizacion = RUC
             End If
          End If
          
          If Len(FA.Autorizacion) >= 13 And FA.Estado_SRI <> "OK" Then
             SQL2 = "UPDATE Facturas " _
                  & "SET RUC_CI = C.CI_RUC, TB = C.TD, Razon_Social = C.Cliente, Direccion_RS = C.Direccion, Telefono_RS = C.Telefono " _
                  & "FROM Facturas As F, Clientes As C " _
                  & "WHERE F.Item = '" & NumEmpresa & "' " _
                  & "AND F.Periodo = '" & Periodo_Contable & "' " _
                  & "AND F.TC = '" & FA.TC & "' " _
                  & "AND F.Serie = '" & FA.Serie & "' " _
                  & "AND F.Factura = " & FA.Factura & " " _
                  & "AND LEN(F.Razon_Social) = 1 " _
                  & "AND F.CodigoC = C.Codigo "
             Ejecutar_SQL_SP SQL2

             SRI_Crear_Clave_Acceso_Facturas FA, True, CBool(CheqClaveAcceso.value), True
             TxtXML = SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion.Clave_De_Acceso)
             TxtXML.Refresh
             RatonNormal
          Else
             Progreso_Final
             MsgBox "Esta Factura ya esta autorizada"
          End If
       Else
          MsgBox "No se puede enviar al SRI autorizar documentos que estan anulados"
       End If
    Else
       RatonNormal
       MsgBox MensajeNoAutorizarCE
    End If
End Sub

Public Sub Volver_Autorizar_Pendientes()
Dim NumFile As Long
Dim LongFA As Long
Dim FechaDeAut As String
Dim FAPend() As Tipo_Facturas
  If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
     Actualizar_Razon_Social MBFecha
     Progreso_Barra.Mensaje_Box = "Procesando Facturas pendientes de autorizar"
     Progreso_Iniciar
     LongFA = 0
     Contador = 0
     TextoImprimio = ""
     FechaDeAut = MBFecha
     
     sSQL = "SELECT CodigoC, Clave_Acceso, Estado_SRI, TC, Fecha, Serie, Factura, Autorizacion " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Serie = '" & DCSerie & "' " _
          & "AND T <> 'A' " _
          & "AND LEN(Autorizacion) = 13 "
     If CheqAutxRangos.value <> 0 And Val(TextFDesde) <= Val(TextFHasta) Then sSQL = sSQL & "AND Factura BETWEEN " & Val(TextFDesde) & " AND " & Val(TextFHasta) & " "
     sSQL = sSQL & "ORDER BY TC,Serie,Factura "
     Select_Adodc AdoFactList, sSQL
     RatonReloj
     '& "AND Estado_SRI <> 'OK' "
     With AdoFactList.Recordset
      If .RecordCount > 0 Then
          LongFA = .RecordCount
          ReDim FAPend(LongFA) As Tipo_Facturas
         'MsgBox sSQL & vbCrLf & .RecordCount
          Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2) + 100
          Factura_Desde = .fields("Factura")
          Do While Not .EOF
             FAPend(Contador).Estado_SRI = "CN"
             FAPend(Contador).TC = .fields("TC")
             FAPend(Contador).Serie = .fields("Serie")
             FAPend(Contador).Fecha = .fields("Fecha")
             FAPend(Contador).Factura = .fields("Factura")
             FAPend(Contador).Autorizacion = .fields("Autorizacion")
             FAPend(Contador).ClaveAcceso = .fields("Clave_Acceso")
             FAPend(Contador).CodigoC = .fields("CodigoC")
             FA.ClaveAcceso = .fields("Clave_Acceso")
             If FA.ClaveAcceso <> Ninguno Then
                Progreso_Barra.Mensaje_Box = "Eliminando Documento XML No. " & FA.ClaveAcceso
                Progreso_Esperar True
                RutaGeneraFile = RutaDocumentos & "\Comprobantes Generados\" & FA.ClaveAcceso & ".xml"
                If Dir$(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
                RutaGeneraFile = RutaDocumentos & "\Comprobantes Firmados\" & FA.ClaveAcceso & ".xml"
                If Dir$(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
                RutaGeneraFile = RutaDocumentos & "\Comprobantes no Autorizados\" & FA.ClaveAcceso & ".xml"
                If Dir$(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
             End If
             Progreso_Barra.Mensaje_Box = Contador & "/" & .RecordCount & " Autorizando Documento No. " & FA.TC & " " & Format(FA.Fecha, "MM/yyyy") & " - " & FA.Serie & "-" & Format(FA.Factura, "000000000")
             Progreso_Esperar
             Contador = Contador + 1
            .MoveNext
          Loop
      Else
          MsgBox "No hay Facturas Pendientes para autorizar"
      End If
     End With
     If LongFA > 0 Then
        Factura_Desde = FAPend(0).Factura
        For Contador = 0 To LongFA - 1
            FA.CodigoC = FAPend(Contador).CodigoC
            FA.ClaveAcceso = FAPend(Contador).ClaveAcceso
            FA.Estado_SRI = FAPend(Contador).Estado_SRI
            FA.TC = FAPend(Contador).TC
            FA.Serie = FAPend(Contador).Serie
            FA.Fecha = FAPend(Contador).Fecha
            FA.Factura = FAPend(Contador).Factura
            FA.Autorizacion = FAPend(Contador).Autorizacion
            Factura_Hasta = FA.Factura
            SRI_Crear_Clave_Acceso_Facturas FA, False, CBool(CheqClaveAcceso.value), True, True
            'MsgBox FA.TC & ": " & FA.Serie & "-" & FA.Factura & vbCrLf & TextoImprimio
            'If SRI_Autorizacion.Estado_SRI <> "OK" Then MsgBox FA.TC & ": " & FA.Serie & "-" & FA.Factura
            Progreso_Barra.Mensaje_Box = Contador & "/" & LongFA & " Autorizando Documento No. " & FA.TC & " " & Format(FA.Fecha, "MM/yyyy") & " - " & FA.Serie & "-" & Format(FA.Factura, "000000000")
            Progreso_Esperar
        Next Contador
     End If
     If TextoImprimio <> "" Then
        RutaGeneraFile = RutaSysBases & "\TEMP\Informe de Errores " & Format(Factura_Desde, "000000000") & "-" & Format(Factura_Hasta, "000000000") & ".txt"
        NumFile = FreeFile
        Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
             Print #NumFile, TextoImprimio;
        Close #NumFile
        MsgBox "ARCHIVO DE INFORME DE ERRORES:" & vbCrLf & vbCrLf & RutaGeneraFile
     End If
     RatonNormal
     Progreso_Final
  Else
     RatonNormal
     MsgBox MensajeNoAutorizarCE
  End If
End Sub

Private Sub CommandButton2_Click()
  Unload ListFact
End Sub

Public Sub Cambia_Fechas_Facturas()
  If ClaveAuxiliar Then
     FechaValida MBFecha
     Mifecha = BuscarFecha(MBFecha)
     MiMes = MesesLetras(Month(MBFecha))
     If DCFact.Text = "" Then DCFact.Text = "0"
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     
     Titulo = "CONFIRMACION DE ACTUALIZACION"
     Mensajes = "¿Realmente desea cambiar a la fecha " & MBFecha.Text & " la emision de " & FA.TC & " Serie " & FA.Serie & vbCrLf _
              & "Autorizacion: " & FA.Autorizacion & " las facturas desde: " & Format(Factura_Desde, "000000000") & " al " & Format(Factura_Hasta, "000000000") & vbCrLf
     If BoxMensaje = vbYes Then
        Control_Procesos "F", "Cambio de Fecha desde " & Factura_Desde & " a la " & Factura_Hasta
        If (Factura_Hasta - Factura_Desde) >= 0 Then
           RatonReloj
           SQL2 = "UPDATE Facturas " _
                & "SET Fecha = #" & Mifecha & "# " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
           Ejecutar_SQL_SP SQL2
           SQL2 = "UPDATE Detalle_Factura " _
                & "SET Fecha = #" & Mifecha & "# " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
           Ejecutar_SQL_SP SQL2
           SQL2 = "UPDATE Trans_Abonos " _
                & "SET Fecha = #" & Mifecha & "# " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
           Ejecutar_SQL_SP SQL2
           RatonNormal
           MsgBox "Proceso Terminado"
        Else
           MsgBox "No se puede procesar el rando de Facturas"
        End If
     End If
  End If
End Sub

Public Sub Cambia_Autorizacion_Facturas()
  If ClaveAdministrador Then
     RatonReloj
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     Control_Procesos "F", "Cambio de Autorizacion desde " & Factura_Desde & " a la " & Factura_Hasta
     If (Factura_Hasta - Factura_Desde) >= 0 Then FCambioFacturas.Show 1
  End If
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     Codigo1 = TrimStrg(SinEspaciosDer(DCArticulo))
     Codigo2 = TrimStrg(MidStrg(DCArticulo, 1, Len(DCArticulo) - Len(Codigo1)))
     Select Case Opcion
       Case 1
            SQL2 = "UPDATE Detalle_Factura " _
                 & "SET Codigo = '" & Codigo1 & "',Producto = '" & Codigo2 & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC = '" & FA.TC & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                 & "AND Factura = " & FA.Factura & " " _
                 & "AND Codigo = '" & CodigoInv & "' "
            Ejecutar_SQL_SP SQL2
            SQL2 = "UPDATE Trans_Kardex " _
                 & "SET Codigo_Inv = '" & Codigo1 & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC = '" & FA.TC & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND Factura = " & FA.Factura & " " _
                 & "AND Codigo_Inv = '" & CodigoInv & "' "
            Ejecutar_SQL_SP SQL2
            Actualiza_Procesado_Kardex Codigo1
            Actualiza_Procesado_Kardex CodigoInv
       Case 2
            SQL2 = "UPDATE Detalle_Factura " _
                 & "SET CodBodega = '" & Codigo1 & "' " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND TC = '" & FA.TC & "' " _
                 & "AND Serie = '" & FA.Serie & "' " _
                 & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                 & "AND Factura = " & FA.Factura & " " _
                 & "AND CodBodega = '" & CodigoInv & "' "
            Ejecutar_SQL_SP SQL2
     End Select
     DCArticulo.Visible = False
     MsgBox "(" & Opcion & ") PROCESO TERMINADO CON EXITO, VUELVA A LISTAR EL DOCUMENTO PARA VERIFICAR"
  End If
  If KeyCode = vbKeyEscape Then
     DCArticulo.Visible = False
  End If
End Sub

Private Sub DCEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFact_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFact_LostFocus()
  FA.Factura = Val(DCFact)
  BuscarFactura
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  FA.Serie = DCSerie
  Listar_Factura_NotaVentas
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
    FA.TC = DCTipo
    If FA.TC = "" Then FA.TC = "FA"
    
   MiTiempo = Time
   MiTiempo1 = Time
    sSQL = "SELECT Serie " _
         & "FROM Facturas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = '" & FA.TC & "' " _
         & "GROUP BY Serie " _
         & "ORDER BY Serie "
    SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
   TiempoTarea = Time
   'MsgBox Format(MiTiempo1 - MiTiempo, "hh:mm:ss.cc") & vbCrLf & Format(TiempoTarea - MiTiempo1, "hh:mm:ss.cc")
    
End Sub

Private Sub DGDetalle_BeforeDelete(Cancel As Integer)
 'If OpcionTab <> 1 Then
  Cancel = True
End Sub

Private Sub DGDetalle_DblClick()
  If OpcionTab = 0 Then
     TxtDetalle.Visible = False
     TxtDetalle.Text = ""
     With AdoDetalle.Recordset
      If .RecordCount Then
          Codigo4 = DGDetalle.Columns(0)
         .MoveFirst
         .Find ("Codigo = '" & Codigo4 & "' ")
          If Not .EOF And .fields("Detalle") <> Ninguno Then
             TxtDetalle.Visible = True
             TxtDetalle.Text = DGDetalle.Columns(1) & ": " & vbCrLf & .fields("Detalle")
             TxtDetalle.SetFocus
          End If
      End If
     End With
  End If
End Sub

Private Sub DGDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    With AdoDetalle.Recordset
     If .RecordCount > 0 Then
         Select Case OpcionTab
           Case 0
                If AltDown And KeyCode = vbKeyF9 Then
                   Opcion = 1
                   If ClaveAdministrador Then
                      CodigoInv = .fields("Codigo")
                      sSQL = "SELECT Producto & '  ' & Codigo_Inv As NombProducto " _
                           & "FROM Catalogo_Productos " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND TC = 'P' " _
                           & "AND Codigo_Inv <> '" & CodigoInv & "' " _
                           & "AND INV <> " & Val(adFalse) & " " _
                           & "ORDER BY Producto, Codigo_Inv "
                      SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "NombProducto"
                     'MsgBox AdoArticulo.Recordset.RecordCount
                      DCArticulo.Visible = True
                      DCArticulo.SetFocus
                   End If
                End If
            
                If AltDown And KeyCode = vbKeyF10 Then
                   Opcion = 2
                   If ClaveAdministrador Then
                      CodigoInv = .fields("CodBodega")
                      sSQL = "SELECT Bodega & '  ' & CodBod As NombProducto " _
                           & "FROM Catalogo_Bodegas " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND CodBod <> '" & CodigoInv & "' " _
                           & "ORDER BY Bodega,CodBod "
                      SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "NombProducto"
                      DCArticulo.Visible = True
                      DCArticulo.SetFocus
                   End If
                End If

                If CtrlDown And KeyCode = vbKeyS Then
                   PosItem = DGDetalle.Columns(DGDetalle.Columns.Count - 1)
                   CodigoP = TrimStrg(InputBox("INGRESE LA SERIE DE ESTE PRODUCTO", "INGRESO DE SERIE"))
                   If Len(CodigoP) > 1 Then
                      sSQL = "UPDATE Detalle_Factura " _
                           & "SET Serie_No = '" & CodigoP & "' " _
                           & "WHERE Factura = " & FA.Factura & " " _
                           & "AND TC = '" & FA.TC & "' " _
                           & "AND Serie = '" & FA.Serie & "' " _
                           & "AND Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND ID = " & PosItem & " "
                      Ejecutar_SQL_SP sSQL
                      MsgBox "Proceso Terminado, vuelva a listar el documento"
                   End If
                End If
                If CtrlDown And KeyCode = vbKeyL Then
                   PosItem = DGDetalle.Columns(DGDetalle.Columns.Count - 1)
                   CodigoInv = DGDetalle.Columns(0)
                   CodigoP = TrimStrg(InputBox("INGRESE EL LOTE DE ESTE PRODUCTO", "INGRESO DE LOTES"))
                   If Len(CodigoP) > 1 Then
                      sSQL = "UPDATE Detalle_Factura " _
                           & "SET Lote_No = '" & CodigoP & "' " _
                           & "WHERE Factura = " & FA.Factura & " " _
                           & "AND TC = '" & FA.TC & "' " _
                           & "AND Serie = '" & FA.Serie & "' " _
                           & "AND Codigo = '" & CodigoInv & "' " _
                           & "AND Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND ID = " & PosItem & " "
                      Ejecutar_SQL_SP sSQL
                      
                      sSQL = "UPDATE Trans_Kardex " _
                           & "SET Lote_No = '" & CodigoP & "' " _
                           & "WHERE Factura = " & FA.Factura & " " _
                           & "AND TC = '" & FA.TC & "' " _
                           & "AND Serie = '" & FA.Serie & "' " _
                           & "AND Codigo_Inv = '" & CodigoInv & "' " _
                           & "AND Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' "
                      Ejecutar_SQL_SP sSQL
                      
                      MsgBox "Proceso Terminado, vuelva a listar el documento"
                   End If
                End If
           Case 1
                If CtrlDown And KeyCode = vbKeyDelete Then
                    Si_No = .fields("C")
                    Total = .fields("Abono")
                    SerieFactura = .fields("Serie")
                    Autorizacion = .fields("Autorizacion")
                    Factura_No = .fields("Factura")
                    CodigoCli = .fields("CodigoC")
                    TipoDoc = .fields("TP")
                    Cta_Aux = .fields("Cta")
                    CodigoB = .fields("Banco")
                    CodigoA = .fields("Cheque")
                    Codigo1 = .fields("Recibo_No")
                    FA.TC = .fields("TP")
                    FA.Serie = .fields("Serie")
                    FA.Factura = .fields("Factura")
                    FA.Autorizacion = .fields("Autorizacion")
                    FA.CodigoC = .fields("CodigoC")
                    FA.Serie_NC = .fields("Serie_NC")
                    FA.Nota_Credito = .fields("Secuencial_NC")
                    FA.Autorizacion_NC = .fields("Autorizacion_NC")
                    ID_Trans = .fields("ID")
                    
                    Titulo = "CONFIRMACION DE ELIMINACION"
                    Mensajes = "¿Realmente desea eliminar el Abono:" & vbCrLf _
                             & "Fecha: " & .fields("Fecha") & vbCrLf _
                             & "Factura No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") & vbCrLf _
                             & "Abono USD " & .fields("Abono") & vbCrLf _
                             & "Detalle: " & CodigoB & " " & CodigoA & vbCrLf
                    If Len(FA.Serie_NC) = 6 And FA.Nota_Credito > 0 And Len(FA.Autorizacion_NC) >= 13 Then
                       Mensajes = Mensajes & "y ademas contiene la siguiente: " & vbCrLf _
                                & "NOTA DE CREDITO: " & FA.Serie_NC & "-" & Format(FA.Nota_Credito, "000000000") & vbCrLf
                    End If
                    If BoxMensaje = vbYes Then
                       Actualiza_Procesado_Kardex_Factura FA
                       sSQL = "UPDATE Facturas " _
                            & "SET Saldo_MN = Saldo_MN + " & Total & ", T = 'P' " _
                            & "WHERE TC = '" & FA.TC & "' " _
                            & "AND Factura = " & FA.Factura & " " _
                            & "AND Serie = '" & FA.Serie & "' " _
                            & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                            & "AND CodigoC = '" & FA.CodigoC & "' " _
                            & "AND Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' "
                       Ejecutar_SQL_SP sSQL
                       
                       sSQL = "UPDATE Detalle_Factura " _
                            & "SET Fecha_NC = #" & BuscarFecha(FechaSistema) & "#," _
                            & "Serie_NC = '" & Ninguno & "'," _
                            & "Autorizacion_NC = '" & Ninguno & "'," _
                            & "Secuencial_NC = 0," _
                            & "Total_IVA_NC = 0," _
                            & "Cantidad_NC = 0," _
                            & "Total_Desc_NC = 0," _
                            & "SubTotal_NC = 0 " _
                            & "WHERE Factura = " & FA.Factura & " " _
                            & "AND TC = '" & FA.TC & "' " _
                            & "AND Serie = '" & FA.Serie & "' " _
                            & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                            & "AND Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' "
                       Ejecutar_SQL_SP sSQL
                       
                       If Len(FA.Serie_NC) = 6 And FA.Nota_Credito > 0 And Len(FA.Autorizacion_NC) >= 13 Then
                          sSQL = "DELETE * " _
                               & "FROM Trans_Abonos " _
                               & "WHERE TP = '" & FA.TC & "' " _
                               & "AND Factura = " & FA.Factura & " " _
                               & "AND Serie = '" & FA.Serie & "' " _
                               & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Contable & "' " _
                               & "AND Serie_NC = '" & FA.Serie_NC & "' " _
                               & "AND Autorizacion_NC = '" & FA.Autorizacion_NC & "' " _
                               & "AND Secuencial_NC = " & FA.Nota_Credito & " "
                          Ejecutar_SQL_SP sSQL
                          
                          sSQL = "DELETE * " _
                               & "FROM Trans_Kardex " _
                               & "WHERE Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Contable & "' " _
                               & "AND TC = '" & FA.TC & "' " _
                               & "AND Serie = '" & FA.Serie & "' " _
                               & "AND Factura = " & FA.Factura & " " _
                               & "AND SUBSTRING(Detalle,1,3) = 'NC:' "
                          Ejecutar_SQL_SP sSQL
                          
                          sSQL = "UPDATE Detalle_Nota_Credito " _
                               & "SET T ='A' " _
                               & "WHERE TC = '" & FA.TC & "' " _
                               & "AND Factura = " & FA.Factura & " " _
                               & "AND Serie_FA = '" & FA.Serie & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Contable & "' " _
                               & "AND Serie = '" & FA.Serie_NC & "' " _
                               & "AND Secuencial = " & FA.Nota_Credito & " "
                          Ejecutar_SQL_SP sSQL
                          
                          sSQL = "UPDATE Trans_Kardex " _
                               & "SET Procesado = 0 " _
                               & "FROM Trans_Kardex As TK, Detalle_Factura As DF " _
                               & "WHERE DF.Item = '" & NumEmpresa & "' " _
                               & "AND DF.Periodo = '" & Periodo_Contable & "' " _
                               & "AND DF.TC = '" & FA.TC & "' " _
                               & "AND DF.Serie = '" & FA.Serie & "' " _
                               & "AND DF.Factura = " & FA.Factura & " " _
                               & "AND DF.Item = TK.Item " _
                               & "AND DF.Periodo = TK.Periodo " _
                               & "AND DF.Codigo = TK.Codigo_Inv "
                          Ejecutar_SQL_SP sSQL
                       Else
                          sSQL = "DELETE * " _
                               & "FROM Trans_Abonos " _
                               & "WHERE TP = '" & FA.TC & "' " _
                               & "AND Factura = " & FA.Factura & " " _
                               & "AND Serie = '" & FA.Serie & "' " _
                               & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Contable & "' " _
                               & "AND Cta = '" & Cta_Aux & "' " _
                               & "AND Banco = '" & CodigoB & "' " _
                               & "AND Cheque = '" & CodigoA & "' " _
                               & "AND CodigoC = '" & CodigoCli & "' " _
                               & "AND Recibo_No = '" & Codigo1 & "' "
                          If Si_No Then
                             sSQL = sSQL & "AND C <> " & adFalse & " "
                          Else
                             sSQL = sSQL & "AND C = " & adFalse & " "
                          End If
                          sSQL = sSQL & "AND ID = " & ID_Trans & " "
                          Ejecutar_SQL_SP sSQL
                       End If
                       
                       sSQL = "DELETE * " _
                            & "FROM Trans_SubCtas " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Fecha <= #" & BuscarFecha(MBFecha) & "# " _
                            & "AND Cta = '" & Cta_Aux & "' " _
                            & "AND Codigo = '" & CodigoCli & "' " _
                            & "AND Comp_No = " & Factura_No & " " _
                            & "AND TP = '.' " _
                            & "AND Numero = 0 " _
                            & "AND T <> 'A' "
                       Ejecutar_SQL_SP sSQL
                                              
                       If Periodo_Contable <> Periodo_Superior Then
                          sSQL = "UPDATE Facturas " _
                               & "SET Saldo_MN = Saldo_MN + " & Total & ", T = 'P' " _
                               & "WHERE TC = '" & TipoDoc & "' " _
                               & "AND Factura = " & Factura_No & " " _
                               & "AND Serie = '" & SerieFactura & "' " _
                               & "AND Autorizacion = '" & Autorizacion & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Superior & "' " _
                               & "AND CodigoC = '" & CodigoCli & "' "
                          Ejecutar_SQL_SP sSQL
                          
                          sSQL = "UPDATE Detalle_Factura " _
                               & "SET Fecha_NC = #" & BuscarFecha(FechaSistema) & "#," _
                               & "Serie_NC = '" & Ninguno & "'," _
                               & "Autorizacion_NC = '" & Ninguno & "'," _
                               & "Secuencial_NC = 0," _
                               & "Total_IVA_NC = 0," _
                               & "Cantidad_NC = 0," _
                               & "Total_Desc_NC = 0," _
                               & "SubTotal_NC = 0 " _
                               & "WHERE Factura = " & FA.Factura & " " _
                               & "AND TC = '" & FA.TC & "' " _
                               & "AND Serie = '" & FA.Serie & "' " _
                               & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Superior & "' "
                          Ejecutar_SQL_SP sSQL
                          
                          sSQL = "DELETE * " _
                               & "FROM Trans_Abonos " _
                               & "WHERE TP = '" & TipoDoc & "' " _
                               & "AND Factura = " & Factura_No & " " _
                               & "AND Serie = '" & SerieFactura & "' " _
                               & "AND Autorizacion = '" & Autorizacion & "' " _
                               & "AND Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Superior & "' " _
                               & "AND Cta = '" & Cta_Aux & "' " _
                               & "AND Banco = '" & CodigoB & "' " _
                               & "AND Cheque = '" & CodigoA & "' " _
                               & "AND CodigoC = '" & CodigoCli & "' " _
                               & "AND Recibo_No = '" & Codigo1 & "' "
                          If Si_No Then
                             sSQL = sSQL & "AND C <> " & adFalse & " "
                          Else
                             sSQL = sSQL & "AND C = " & adFalse & " "
                          End If
                          sSQL = sSQL & "AND ID = " & ID_Trans & " "
                          Ejecutar_SQL_SP sSQL
                      
                          sSQL = "DELETE * " _
                               & "FROM Trans_Kardex " _
                               & "WHERE Item = '" & NumEmpresa & "' " _
                               & "AND Periodo = '" & Periodo_Superior & "' " _
                               & "AND TC = '" & FA.TC & "' " _
                               & "AND Serie = '" & FA.Serie & "' " _
                               & "AND Factura = " & FA.Factura & " " _
                               & "AND SUBSTRING(Detalle,1,3) = 'NC:' "
                          Ejecutar_SQL_SP sSQL
                       End If
''                       FA.Factura = Val(DCFact)
''                       BuscarFactura
                       SQL2 = "SELECT C,T,Fecha,Banco,Cheque,Abono,Serie,Factura,Autorizacion,Protestado,CodigoC,Cta_CxP,Cta,Tipo_Cta,Fecha_Aut_NC,Serie_NC,Secuencial_NC," _
                            & "Autorizacion_NC,Clave_Acceso_NC,TP,Recibo_No,Comprobante,Estado_SRI_NC,Hora_Aut_NC,Periodo,Item,CodigoU,Cod_Ejec,ID " _
                            & "FROM Trans_Abonos " _
                            & "WHERE Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' " _
                            & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                            & "AND TP = '" & FA.TC & "' " _
                            & "AND Serie = '" & FA.Serie & "' " _
                            & "AND Factura = " & FA.Factura & " " _
                            & "ORDER BY TP,Fecha,Cta,Cta_CxP,Abono,Banco,Cheque "
                       Select_Adodc_Grid DGDetalle, AdoDetalle, SQL2
                       MsgBox "Proceso realizado con exito"
                    End If
                End If
                If CtrlDown And KeyCode = vbKeyP Then
                    TA.Recibi_de = FA.Cliente
                    TA.Abono = .fields("Abono")
                    TA.Tipo_Cta = "TJ"
                    Imprimir_FA_NV_TJ TA
                End If
                
                If CtrlDown And KeyCode = vbKeyA Then Volver_Autorizar_NC True
                
                If AltDown And KeyCode = vbKeyN Then Actualizar_NC_Kardex
                
                If KeyCode = vbKeyReturn Then
                   AdoDetalle.Recordset.MoveNext
                   If AdoDetalle.Recordset.EOF Then AdoDetalle.Recordset.MoveFirst
                End If
         End Select
   End If
  End With
End Sub

Private Sub Form_Activate()
   
   Una_Vez = True
   LstBox.Clear
   LstBox.AddItem "<Alt><F9>: Cambiar Codigo de Articulo"
   LstBox.AddItem "<Alt><F10>: Cambiar Codigo de Bodega"
   LstBox.AddItem "<Ctrl><P>: Imprime Recibo de Pago con Tarjeta de Credito"
   LstBox.AddItem "<Ctrl><S>: Cambiar la serie del Producto"
   LstBox.AddItem "<Ctrl><Delete>: Eliminar Abonos "
   
   FA.Fecha_Corte = FechaSistema
   FA.Serie = Ninguno
  'MsgBox FA.TC & vbCrLf & FA.Serie & vbCrLf & FA.Factura
   Encerar_Factura FA
   
   ListFact.WindowState = vbMaximized
   
   SSTabDetalle.Tab = 0
   SSTabDetalle.width = MDI_X_Max - 100
   
   DGDetalle.width = MDI_X_Max - 100
   DGDetalle.Height = MDI_Y_Max - DGDetalle.Top - 950
   
   TxtXML.width = (MDI_X_Max / 2) - 150
   TxtXML.Left = (MDI_X_Max / 2) + 50
   
   LstBox.width = SSTabDetalle.width - LstBox.Left
   Label8.width = (MDI_X_Max / 2) - Label8.Left
   LabelTransp.width = SSTabDetalle.width - LabelTransp.Left
   TxtObs.width = SSTabDetalle.width - TxtObs.Left

   CommandButton2.Top = DGDetalle.Top + DGDetalle.Height
   FrmTotales.Top = DGDetalle.Top + DGDetalle.Height + 40

   If NombreUsuario = "Administrador de Red" Then
      TBarFactura.buttons("Cambio_Emision_Facturas").Enabled = True
      TBarFactura.buttons("Cambio_Vencimiento_Facturas").Enabled = True
      TBarFactura.buttons("Cambia_Autorizacion_Facturas").Enabled = True
      TBarFactura.buttons("Cambia_Numero_de_Facturas").Enabled = True
      TBarFactura.buttons("Reprocesar_Saldos_Facturas").Enabled = True
      TBarFactura.buttons("Eliminar_Facturas").Enabled = True
      TBarFactura.buttons("Revertir_Facturas").Enabled = True
      TBarFactura.buttons("Actualizar_Representantes").Enabled = True
      TBarFactura.buttons("Liberar_FA_SRI").Enabled = True
      TBarFactura.buttons("Ejecutivo").Enabled = True
      'TBarFactura.buttons("Kardex").Enabled = True
   Else
      'TBarFactura.buttons("Cambio_Emision_Facturas").Enabled = False
      TBarFactura.buttons("Cambio_Vencimiento_Facturas").Enabled = False
      TBarFactura.buttons("Cambia_Autorizacion_Facturas").Enabled = False
      TBarFactura.buttons("Cambia_Numero_de_Facturas").Enabled = False
      TBarFactura.buttons("Reprocesar_Saldos_Facturas").Enabled = False
      TBarFactura.buttons("Eliminar_Facturas").Enabled = False
      TBarFactura.buttons("Revertir_Facturas").Enabled = False
      TBarFactura.buttons("Actualizar_Representantes").Enabled = False
      TBarFactura.buttons("Liberar_FA_SRI").Enabled = False
      TBarFactura.buttons("Ejecutivo").Enabled = False
      'TBarFactura.buttons("Kardex").Enabled = False
   End If
   
   sSQL = "SELECT CR.Codigo,C.Cliente,C.CI_RUC,CR.Porc_Com " _
        & "FROM Catalogo_Rol_Pagos As CR INNER JOIN Clientes As C " _
        & "ON CR.Codigo = C.Codigo " _
        & "WHERE CR.Item = '" & NumEmpresa & "' " _
        & "AND CR.Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY C.Cliente "
   SelectDB_Combo DCEjecutivo, AdoEjecutivo, sSQL, "Cliente"
   
   sSQL = "SELECT TC " _
        & "FROM Facturas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "GROUP BY TC " _
        & "ORDER BY TC "
   SelectDB_Combo DCTipo, AdoTipo, sSQL, "TC"
   If AdoTipo.Recordset.RecordCount > 0 Then
      DCTipo.Text = "FA"
      FA.TC = DCTipo
      FA.Serie = "001001"
      FA.Factura = 0
     'BuscarFactura
      TextFDesde = DCFact
      TextFHasta = DCFact
     
     'MsgBox Modulo
      RatonNormal
      MBFecha = FechaSistema
      DCTipo.SetFocus
   Else
      RatonNormal
      MsgBox "ESTA EMPRESA NO A EMPEZADO A PROCESAR " & vbCrLf & vbCrLf & "FACTURAS/NOTAS DE VENTAS"
      Unload ListFact
   End If
End Sub

Private Sub Form_Deactivate()
  ListFact.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
  'Abriendo bases relacionadas
   ConectarAdodc AdoFactura
   ConectarAdodc AdoDetalle
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoDiarioCaja
  'Bases de Listar Factura
   ConectarAdodc AdoTipo
   ConectarAdodc AdoSerie
   ConectarAdodc AdoFactList
   ConectarAdodc AdoArticulo
   ConectarAdodc AdoEjecutivo
   ConectarAdodc AdoAutorizacion
   
  'SRI_Obtener_Datos_Comprobantes_Electronicos
End Sub

Public Sub BuscarFactura()
'Dim CSQL1, CSQL2, CSQL3, CSQL4, CSQL5, CSQL6 As String
  RatonReloj
  TxtXML = ""
  DGDetalle.Visible = False
  DGDetalle.BackColor = &H80000005
 'Volvemos a recalcular los totales de la factura
  With AdoFactList.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & FA.Factura & " ")
       If Not .EOF Then FA.Autorizacion = .fields("Autorizacion")
   End If
  End With
  Leer_Datos_FA_NV FA
 'Procesamos Factura
  If FA.Si_Existe_Doc Then
     Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
  Else
     DGDetalle.Visible = True
     RatonNormal
     MsgBox "Esta Factura no existe."
     FA.Factura = 0
     DCTipo.SetFocus
  End If

    'Consultamos el detalle de la factura
     SQL2 = "SELECT DF.Codigo,DF.Producto,DF.Cantidad,DF.Precio,DF.Total,DF.Total_Desc,DF.Total_Desc2,DF.Total_IVA," _
          & "ROUND(((DF.Total-(DF.Total_Desc+DF.Total_Desc2))+DF.Total_IVA),2,0) As Valor_Total,DF.Mes,DF.Ticket," _
          & "DF.Serie,DF.Factura,DF.Autorizacion,CP.Detalle,CP.Cta_Ventas,CP.Reg_Sanitario,CP.Marca,Lote_No, DF.Modelo, " _
          & "DF.Procedencia, DF.Serie_No, DF.CodigoC, Cantidad_NC, SubTotal_NC,DF.CodMarca,DF.CodBodega," _
          & "DF.Tonelaje,Total_Desc_NC,Total_IVA_NC,DF.Periodo,DF.Codigo_Barra,DF.ID " _
          & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
          & "WHERE DF.Item = '" & NumEmpresa & "' " _
          & "AND DF.Periodo = '" & Periodo_Contable & "' " _
          & "AND DF.TC = '" & FA.TC & "' " _
          & "AND DF.Serie = '" & FA.Serie & "' " _
          & "AND DF.Autorizacion = '" & FA.Autorizacion & "' " _
          & "AND DF.Factura = " & FA.Factura & " " _
          & "AND DF.Periodo = CP.Periodo " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Codigo = CP.Codigo_Inv " _
          & "ORDER BY CP.Cta_Ventas,DF.Codigo,DF.ID "
     SQLDec = "Precio " & CStr(Dec_PVP) & "|Total 2|Total_IVA 4|."
     Select_Adodc_Grid DGDetalle, AdoDetalle, SQL2, SQLDec
     DGDetalle.Visible = True
     FinBucle = True
    'Recolectamos los item de la factura a buscar
     LabelEstado.Caption = FA.T
     Label7.Caption = FA.Grupo
     LabelFechaPe.Caption = FA.Fecha
     FechaComp = FA.Fecha
     LabelCodigo.Caption = FA.CodigoC
     LabelCliente.Caption = FA.Cliente
     Label8.Caption = "Razon Social: " & FA.Razon_Social & ", CI/RUC: " & FA.CI_RUC & vbCrLf _
                    & "Dirección: " & FA.DireccionC & ", Teléfono: " & FA.TelefonoC & vbCrLf _
                    & "Emails: " & FA.EmailC & "; " & FA.EmailR & vbCrLf _
                    & "Forma de pago: " & FA.Forma_Pago & vbCrLf _
                    & "Elaborado por: " & FA.Digitador & " (" & FA.Hora & ")"
     LabelVendedor.Caption = " Ejecutivo: " & FA.Ejecutivo_Venta
     DireccionGuia = FA.Comercial
     TxtAutorizacion = FA.Autorizacion
     TxtClaveAcceso = FA.ClaveAcceso
     TxtObs = FA.Observacion
     LabelTransp.Caption = FA.Nota
     Label15.Caption = DireccionGuia
    'LabelFormaPa.Caption = .Fields("Forma_Pago")
     LabelServicio.Caption = Format$(FA.Servicio, "#,##0.00")
     LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
     LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
     LabelDesc.Caption = Format$(FA.Descuento + FA.Descuento2, "#,##0.00")
     
     LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
     LabelServicio.Caption = Format$(FA.Servicio, "#,##0.00")
     LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
     LabelSaldoAct.Caption = Format$(FA.Saldo_MN, "#,##0.00")
     Select Case LabelEstado.Caption
       Case Anulado:            LabelEstado.Caption = "Anulada"
       Case Pendiente, Normal:  LabelEstado.Caption = "Pendiente"
       Case Cancelado:          LabelEstado.Caption = "Cancelada"
       Case Else:               LabelEstado.Caption = "No existe"
     End Select
    'Consultamos los pagos Interes de Tarjetas y Abonos de Bancos con efectivo
    'Procesamos el Saldo de la Factura
  
  CheqClaveAcceso.Caption = "Clave de Accceso: " & FA.TC & " "
  TxtXML = SRI_Leer_Comprobantes_no_Autorizados(SRI_Autorizacion.Clave_De_Acceso)
  TxtXML.Refresh
  SSTabDetalle.Tab = 0
End Sub

Private Sub LstBox_KeyDown(KeyCode As Integer, Shift As Integer)
Keys_Especiales Shift
    If KeyCode = vbKeyEscape Then LstBox.Visible = False
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub SSTabDetalle_Click(PreviousTab As Integer)
Dim AdoAuxDB As ADODB.Recordset
    RatonReloj
    DGDetalle.AllowDelete = False
    DGDetalle.AllowUpdate = False
    DGDetalle.Visible = False
'    TxtXML.Visible = False
    FrmTotalAsiento.Visible = False
    DGDetalle.Height = MDI_Y_Max - DGDetalle.Top - 950
    OpcionTab = SSTabDetalle.Tab
    Select Case SSTabDetalle.Tab
      Case 0: 'DETALLE DE FACTURA
           DGDetalle.BackColor = &H80000005
           SQL2 = "SELECT DF.Codigo,DF.Producto,DF.Cantidad,DF.Precio,DF.Total,DF.Total_Desc,DF.Total_Desc2,DF.Total_IVA," _
                & "ROUND(((DF.Total-(DF.Total_Desc+DF.Total_Desc2))+DF.Total_IVA),2,0) As Valor_Total,DF.Mes,DF.Ticket," _
                & "DF.Serie,DF.Factura,DF.Autorizacion,CP.Detalle,CP.Cta_Ventas,CP.Reg_Sanitario,CP.Marca,Lote_No, DF.Modelo, " _
                & "DF.Procedencia, DF.Serie_No, DF.CodigoC, Cantidad_NC, SubTotal_NC,DF.CodMarca,DF.CodBodega," _
                & "DF.Tonelaje,Total_Desc_NC,Total_IVA_NC,DF.Periodo,DF.Codigo_Barra,DF.ID " _
                & "FROM Detalle_Factura As DF,Catalogo_Productos As CP " _
                & "WHERE DF.Item = '" & NumEmpresa & "' " _
                & "AND DF.Periodo = '" & Periodo_Contable & "' " _
                & "AND DF.TC = '" & FA.TC & "' " _
                & "AND DF.Serie = '" & FA.Serie & "' " _
                & "AND DF.Autorizacion = '" & FA.Autorizacion & "' " _
                & "AND DF.Factura = " & FA.Factura & " " _
                & "AND DF.Periodo = CP.Periodo " _
                & "AND DF.Item = CP.Item " _
                & "AND DF.Codigo = CP.Codigo_Inv " _
                & "ORDER BY CP.Cta_Ventas,DF.Codigo,DF.ID "
           SQLDec = "Precio " & CStr(Dec_PVP) & "|Total 2|Total_IVA 4|."
           Select_Adodc_Grid DGDetalle, AdoDetalle, SQL2, SQLDec
           DGDetalle.Visible = True
      Case 1: 'ABONOS DE LA FACTURA
           DGDetalle.BackColor = &H80C0FF
           DGDetalle.AllowDelete = True
           DGDetalle.AllowUpdate = True
           FA.Total_Abonos = 0
           FA.SubTotal_NC = 0
           FA.Total_IVA_NC = 0
           SQL2 = "SELECT C,T,Fecha,Banco,Cheque,Abono,Serie,Factura,Autorizacion,Protestado,CodigoC,Cta_CxP,Cta,Tipo_Cta,Fecha_Aut_NC,Serie_NC,Secuencial_NC," _
                & "Autorizacion_NC,Clave_Acceso_NC,TP,Recibo_No,Comprobante,Estado_SRI_NC,Hora_Aut_NC,Periodo,Item,CodigoU,Cod_Ejec,ID " _
                & "FROM Trans_Abonos " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP IN ('" & FA.TC & "','TJ') " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "ORDER BY TP,Fecha,Cta,Cta_CxP,Abono,Banco,Cheque "
           Select_Adodc_Grid DGDetalle, AdoDetalle, SQL2
           If AdoDetalle.Recordset.RecordCount > 0 Then
             'Len(AdoAbonos.Recordset.Fields("Clave_Acceso_NC")) >= 13 And
              Do While Not AdoDetalle.Recordset.EOF
                 If AdoDetalle.Recordset.fields("TP") <> "TJ" Then
                    FA.Total_Abonos = FA.Total_Abonos + AdoDetalle.Recordset.fields("Abono")
                    If AdoDetalle.Recordset.fields("Banco") = "NOTA DE CREDITO" Then
                       FA.Porc_NC = FA.Porc_IVA
                       FA.Fecha_NC = AdoDetalle.Recordset.fields("Fecha")
                       FA.Fecha_Aut_NC = AdoDetalle.Recordset.fields("Fecha_Aut_NC")
                       FA.Serie_NC = AdoDetalle.Recordset.fields("Serie_NC")
                       FA.Nota_Credito = AdoDetalle.Recordset.fields("Secuencial_NC")
                       FA.Autorizacion_NC = AdoDetalle.Recordset.fields("Autorizacion_NC")
                       FA.ClaveAcceso_NC = AdoDetalle.Recordset.fields("Clave_Acceso_NC")
                       If AdoDetalle.Recordset.fields("Cheque") = "VENTAS" Then
                          FA.SubTotal_NC = FA.SubTotal_NC + AdoDetalle.Recordset.fields("Abono")
                       Else
                          FA.Total_IVA_NC = FA.Total_IVA_NC + AdoDetalle.Recordset.fields("Abono")
                       End If
                    End If
                 End If
                 AdoDetalle.Recordset.MoveNext
              Loop
           End If
           DGDetalle.Visible = True
      Case 2: 'GUIA DE REMISION
           DGDetalle.BackColor = &HC0FFFF
           SQL2 = "SELECT Serie_GR, Remision, Clave_Acceso_GR, Autorizacion_GR, Fecha, CodigoC, Comercial, CIRUC_Comercial, " _
                & "Entrega, CIRUC_Entrega, CiudadGRI, CiudadGRF, Placa_Vehiculo, FechaGRE, FechaGRI, FechaGRF, Pedido, Zona, " _
                & "Hora_Aut_GR, Estado_SRI_GR, Error_FA_SRI, Fecha_Aut_GR, TC, Serie, Factura, Autorizacion, Lugar_Entrega, " _
                & "Periodo, Item " _
                & "FROM Facturas_Auxiliares " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Remision > 0 " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' "
           Select_Adodc_Grid DGDetalle, AdoDetalle, SQL2
           DGDetalle.Visible = True
      Case 3: 'CONTABILIZACION
           sSQL = "DELETE * " _
                & "FROM Asiento " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND T_No IN(253,254) " _
                & "AND CodigoU = '" & CodigoUsuario & "' "
           Ejecutar_SQL_SP sSQL
           Trans_No = 253
           Insertar_Ctas_Cierre_SP "CXC", 1
           DGDetalle.BackColor = &HC0FFC0
           DGDetalle.Visible = False
           DGDetalle.Height = MDI_Y_Max - DGDetalle.Top - 1550
           FrmTotalAsiento.Top = DGDetalle.Top + DGDetalle.Height + 40
           
           sSQL = "SELECT Cta_Venta,SUM(Total) As TTotal,SUM(Total_Desc) As TTotal_Desc,SUM(Total_Desc2) As TTotal_Desc2,SUM(Total_IVA) As TTotal_IVA " _
                & "FROM Detalle_Factura " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                & "GROUP BY Cta_Venta " _
                & "ORDER BY Cta_Venta "
           Select_AdoDB AdoAuxDB, sSQL
           With AdoAuxDB
            If .RecordCount > 0 Then
                Do While Not .EOF
                   Valor = .fields("TTotal") - .fields("TTotal_Desc") - .fields("TTotal_Desc2") + .fields("TTotal_IVA")
                   Insertar_Ctas_Cierre_SP FA.Cta_CxP, Valor
                   Insertar_Ctas_Cierre_SP .fields("Cta_Venta"), -.fields("TTotal")
                   Insertar_Ctas_Cierre_SP Cta_Desc, .fields("TTotal_Desc")
                   Insertar_Ctas_Cierre_SP Cta_Desc2, .fields("TTotal_Desc2")
                   Insertar_Ctas_Cierre_SP Cta_IVA, -.fields("TTotal_IVA")
                  .MoveNext
                Loop
            End If
           End With
           AdoAuxDB.Close
           Trans_No = 254
           Insertar_Ctas_Cierre_SP "ABONO", 1
           sSQL = "SELECT Cta, Cta_CxP, SUM(Abono) As TAbono " _
                & "FROM Trans_Abonos " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                & "GROUP BY Cta, Cta_CxP "
           Select_AdoDB AdoAuxDB, sSQL
           With AdoAuxDB
            If .RecordCount > 0 Then
                Do While Not .EOF
                   Insertar_Ctas_Cierre_SP .fields("Cta"), .fields("TAbono")
                   Insertar_Ctas_Cierre_SP .fields("Cta_CxP"), -.fields("TAbono")
                  .MoveNext
                Loop
            End If
           End With
           AdoAuxDB.Close
           RatonReloj
           Debe = 0
           Haber = 0
           sSQL = "SELECT * " _
                & "FROM Asiento " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND T_No IN(253,254) " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "ORDER BY T_No, A_No, DEBE DESC,HABER "
           Select_Adodc_Grid DGDetalle, AdoDetalle, sSQL
           With AdoDetalle.Recordset
            If .RecordCount > 0 Then
                Do While Not .EOF
                   Debe = Debe + .fields("Debe")
                   Haber = Haber + .fields("Haber")
                  .MoveNext
                Loop
            End If
           End With
           LabelDebe.Caption = Format$(Debe, "#,##0.00")
           LabelHaber.Caption = Format$(Haber, "#,##0.00")
           LblDiferencia.Caption = Format$(Debe - Haber, "#,##0.00")
           FrmTotalAsiento.Visible = True
           DGDetalle.Visible = True
    End Select
    RatonNormal
End Sub

Private Sub TBarFactura_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  'MsgBox ButtonMenu.key
   Select Case ButtonMenu.key
     Case "PRN_Individual"
          Control_Procesos "I", "Reimpresion de Facturas"
          Imprimir_Facturas FA
          Facturas_Impresas FA
     Case "PRN_En_Bloque"
          Control_Procesos "I", "Reimpresion de Facturas"
          Impresion_Bloque
          Facturas_Impresas FA
     Case "PRN_PV"
          Control_Procesos "I", "Reimpresion de Facturas"
          Imprimir_Punto_Venta FA
''          If Grafico_PV Then
''             Imprimir_Punto_Venta_Grafico FA
''          Else
''             Imprimir_Punto_Venta FA
''          End If
          Facturas_Impresas FA
     Case "PRN_NC"
          Imprimir_NC
          Control_Procesos "I", "Impresion de Nota de Credito No. " & FA.Factura
     Case "PRN_Recibos"
          Imprimir_Recibos CheqSinCodigo.value
     Case "PRN_Guia_R"
          Mensajes = "IMPRIMIR GUIA DE REMISION"
          Titulo = "IMPRESION"
          If BoxMensaje = vbYes Then Imprimir_Guia_Remision AdoFactura, AdoDetalle, FA
    '=================================================================================
    'Envio de Emails de Comprobantes Electronicos
    '=================================================================================
     Case "Mail_FA"
          SRI_Autorizacion.Tipo_Doc_SRI = "FA"
          SRI_Enviar_Mails FA, SRI_Autorizacion
     Case "Mail_NC"
          SRI_Autorizacion.Tipo_Doc_SRI = "NC"
          SRI_Enviar_Mails FA, SRI_Autorizacion
     Case "Mail_GR"
          SRI_Autorizacion.Tipo_Doc_SRI = "GR"
          SRI_Enviar_Mails FA, SRI_Autorizacion
    '=================================================================================
    'Descarga de XML de Comprobantes Electronicos
    '=================================================================================
     Case "XML_FA"
          SRI_Generar_XML_Firmado FA.ClaveAcceso
          MsgBox "Documento Electronico:" & vbCrLf & FA.ClaveAcceso & ".xml" & vbCrLf & "Descargado en: " & RutaSysBases & "\TEMP\"
     Case "XML_LC"
          SRI_Generar_XML_Firmado FA.ClaveAcceso_LC
          MsgBox "Documento Electronico:" & vbCrLf & FA.ClaveAcceso_LC & ".xml" & vbCrLf & "Descargado en: " & RutaSysBases & "\TEMP\"
     Case "XML_GR"
          SRI_Generar_XML_Firmado FA.ClaveAcceso_GR
          MsgBox "Documento Electronico:" & vbCrLf & FA.ClaveAcceso_GR & ".xml" & vbCrLf & "Descargado en: " & RutaSysBases & "\TEMP\"
    '=================================================================================
    'Envio de PDF de Comprobantes Electronicos
    '=================================================================================
     Case "PDF_OP"
          SRI_Generar_PDF_FA FA, True
     Case "PDF_FA"
          SRI_Generar_PDF_FA FA, True
     Case "PDF_LC"
          SRI_Generar_PDF_FA FA, True
    '=================================================================================
    'Generar_XML_Facturas FA
    '=================================================================================
     Case "PDF_NC"
          SRI_Generar_PDF_NC FA, True
     Case "PDF_GR"
          SRI_Generar_PDF_GR FA, True
     Case "PDF_DO"
'''          SetPrinters.Show 1
'''          If PonImpresoraDefecto(SetNombrePRN) Then
          Generar_PDF_Donacion FA, True
    '=================================================================================
    ' AUTORIZACION DE COMPROBANTES ELECTRONICOS PENDIENTE CON EL SRI
    '=================================================================================
     Case "SRIFactAct"
          Volver_Autorizar
     Case "SRIAutNC"
          Volver_Autorizar_NC
     Case "GuiaR"
          Volver_Autorizar_GR
     Case "SRIFactPend"
          Volver_Autorizar_Pendientes
   End Select
End Sub

Private Sub TextFDesde_GotFocus()
   MarcarTexto TextFDesde
End Sub

Private Sub TextFDesde_LostFocus()
   TextoValido TextFDesde
End Sub

Private Sub TextFHasta_GotFocus()
  MarcarTexto TextFHasta
End Sub

Private Sub TextFHasta_LostFocus()
  TextoValido TextFHasta
End Sub

Private Sub TBarFactura_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim NivelNo1
Dim TipoProc1
Dim TipoFacturas1
Dim TempStrg As String
 Factura_No = FA.Factura
 TipoFactura = FA.TC
 CodigoCliente = FA.CodigoC
 
 If Button.key <> "Salir" And Una_Vez Then
    Actualizar_Todas_Razon_Social_SP MBFecha
    Una_Vez = False
 End If
'MsgBox Button.key
 Select Case Button.key
   Case "Salir"
        Unload Me
  'Barra de Separacion
   Case "Cambio_Emision_Facturas"
        Cambia_Fechas_Facturas
   Case "Cambio_Vencimiento_Facturas"
        Cambia_Vencimiento_Facturas
   Case "Cambia_Autorizacion_Facturas"
        Cambia_Autorizacion_Facturas
   Case "Cambia_Numero_de_Facturas"
        Cambia_Numero_de_Facturas
   Case "Reprocesar_Saldos_Facturas"
        If ClaveAuxiliar Then
           DGDetalle.Visible = False
           Actualizar_Saldo_De_Facturas_SP FA.TC, FA.Serie, TextFDesde, TextFHasta, MBFecha
           DGDetalle.Visible = True
           FInfoError.Show
        End If
   Case "Eliminar_Facturas"
        Eliminar_Facturas
   Case "Revertir_Facturas"
        Revertir_Facturas
   Case "Actualizar_Representantes"
        Actualizar_Representantes
  'Barra de Separacion
   Case "Anular_Factura"
        sSQL = "SELECT T " _
             & "FROM Facturas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
             & "AND Factura = " & FA.Factura & " "
        Select_Adodc AdoFactura, sSQL
        With AdoFactura.Recordset
         If .RecordCount > 0 Then
             If .fields("T") = "A" Then
                 MsgBox "Esta Factura ya esta anulada"
             ElseIf ClaveAuxiliar Then
                 Si_No = False
                 FAnulacion.Show
             End If
         End If
        End With
   Case "Anular_en_masa"
        Anular_en_masa
''   Case "Nota_Credito"
''        sSQL = "SELECT * " _
''             & "FROM Facturas " _
''             & "WHERE Item = '" & NumEmpresa & "' " _
''             & "AND Periodo = '" & Periodo_Contable & "' " _
''             & "AND TC = '" & FA.TC & "' " _
''             & "AND Serie = '" & FA.Serie & "' " _
''             & "AND Autorizacion = '" & FA.Autorizacion & "' " _
''             & "AND Factura = " & FA.Factura & " "
''        Select_Adodc AdoFactura, sSQL
''        With AdoFactura.Recordset
''         If .RecordCount > 0 Then
''             If .Fields("T") = "A" Then
''                 MsgBox "Esta Factura ya esta anulada, no se puede emitir una Nota de Credito"
''             ElseIf ClaveAuxiliar Then
''                 Si_No = True
''                 NC.Serie = FA.Serie
''                 FAnulacion.Show
''             End If
''         End If
''        End With
''         MsgBox "Nada"
   Case "CambioCliente"
        RatonReloj
        Factura_No = Val(DCFact.Text)
        TipoDoc = DCTipo.Text
        FCambioCliente.Show
   Case "OP"
        Orden_No = Val(InputBox("Imprimir el Detalle" & vbCrLf & "de la Orden No. ", "IMPRESION DE ORDEN DE TRABAJO", "0"))
        sSQL = "SELECT Fecha,Producto,Cantidad,Precio,A,L,S " _
             & "FROM Trans_Ticket " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Ticket = " & Orden_No & " " _
             & "AND TC = 'OP' " _
             & "ORDER BY Producto "
        Select_Adodc AdoDetAcomp, sSQL
        With AdoDetAcomp.Recordset
         If .RecordCount > 0 Then
             Mensajes = "Imprimir Orden de Trabajo"
             Titulo = "IMPRESION"
             MensajeEncabData = "LISTA DE ORDEN DE TRABAJO No. " & Format$(Orden_No, "000000")
             SQLMsg1 = ""   '"Cliente: " & DCCliente
             Cuadricula = True
             If BoxMensaje = vbYes Then ImprimirAdo AdoDetAcomp, True, 1, 8, True
         End If
        End With
   Case "Liberar_FA_SRI"
        Liberar_FA_SRI
   Case "Ejecutivo"
        FrmEjecutivo.Top = TBarFactura.Height + LabelFechaPe.Top
        FrmEjecutivo.Left = LabelFechaPe.Left
        FrmEjecutivo.Refresh
        FrmEjecutivo.Visible = True
        DCEjecutivo.SetFocus
   Case "Kardex"
        Actualizar_Kardex
   Case "Excel"
        If AdoDetalle.Recordset.RecordCount > 0 Then
           DGDetalle.Visible = False
           GenerarDataTexto ListFact, AdoDetalle
           DGDetalle.Visible = True
        Else
           MsgBox "No existe datos para exportar"
        End If
   Case "Ayuda"
        LstBox.Visible = True
        LstBox.SetFocus
 End Select
 RatonNormal
End Sub

Private Sub TxtAutorizacion_GotFocus()
   MarcarTexto TxtAutorizacion
End Sub

Private Sub TxtAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Secuencial As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyA Then
     If CFechaLong(LabelFechaPe) <= CFechaLong(Fecha_CE) Then
        Secuencial = "INGRESE LA AUTORIZACION DEL" & vbCrLf _
                   & "DOCUMENTO " & FA.TC & "-" & FA.Serie & "-" & Format$(FA.Factura, "000000000") & vbCrLf _
                   & "Autorizacion: " & FA.Autorizacion & ":"
        Autorizacion = InputBox(Secuencial, "CAMBIO DE AUTORIZACION", FA.Autorizacion)
        If IsNumeric(Autorizacion) And Len(Autorizacion) >= 3 Then
           SQL1 = "UPDATE Facturas " _
                & "SET Autorizacion = '" & Autorizacion & "' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' "
           Ejecutar_SQL_SP SQL1
           
           SQL1 = "UPDATE Detalle_Factura " _
                & "SET Autorizacion = '" & Autorizacion & "' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' "
           Ejecutar_SQL_SP SQL1
           
           SQL1 = "UPDATE Trans_Abonos " _
                & "SET Autorizacion = '" & Autorizacion & "' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura = " & FA.Factura & " " _
                & "AND Autorizacion = '" & FA.Autorizacion & "' "
           Ejecutar_SQL_SP SQL1
           MsgBox "Proceso terminado"
        End If
     Else
        RatonNormal
        MsgBox MensajeNoAutorizarCE
     End If
  End If
End Sub

Private Sub TxtClaveAcceso_KeyDown(KeyCode As Integer, Shift As Integer)
Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyR Then ReCalcular_Totales_Factura FA
    If CtrlDown And KeyCode = vbKeyS And Len(FA.Autorizacion) >= 13 Then Volver_Autorizar True
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then TxtDetalle.Visible = False
End Sub

Private Sub TxtObs_GotFocus()
  MarcarTexto TxtObs
End Sub

Private Sub TxtObs_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TextoObservacion As String
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyO Then
    Mensajes = "Poner Observacion en la Factura No. " & FA.Serie & "-" & FA.Factura & vbCrLf & vbCrLf _
             & "Autorizacion: " & FA.Autorizacion
    Titulo = "ACTUALIZACION EN LA FACTURA"
    TextoObservacion = TrimStrg(MidStrg(InputBox(Mensajes, Titulo, TxtObs), 1, 55))
    If Len(TextoObservacion) > 1 Then
       SQL2 = "UPDATE Facturas " _
            & "SET Observacion = '" & TextoObservacion & "' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND TC = '" & FA.TC & "' " _
            & "AND Serie = '" & FA.Serie & "' " _
            & "AND Autorizacion = '" & FA.Autorizacion & "' " _
            & "AND Factura = " & FA.Factura & " "
       Ejecutar_SQL_SP SQL2
       MsgBox "Proceso Exitoso"
    End If
  End If
End Sub

Private Sub TxtObs_LostFocus()
  TextoValido TxtObs, , True
End Sub

Public Sub Impresion_Bloque()
  TextoValido TxtObs, , True
  TextoValido TextFDesde
  TextoValido TextFHasta
  
  If DCFact = "" Then DCFact = "0"
  If Val(TextFDesde) <= 0 Then TextFDesde = DCFact
  If Val(TextFHasta) <= 0 Then TextFHasta = DCFact
  FA.Tipo_PRN = "FM"
  If FA.TC = "OP" Then FA.Tipo_PRN = "OP"
  FA.Desde = Val(TextFDesde)
  FA.Hasta = Val(TextFHasta)
  FA.Factura = Val(DCFact)
  FA.TC = DCTipo
  FA.Serie = DCSerie
  If Len(TxtObs) > 1 Then
     SQL2 = "UPDATE Facturas " _
          & "SET Observacion = '" & TxtObs & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " "
     Ejecutar_SQL_SP SQL2
  End If
  If (FA.Hasta - FA.Desde) >= 0 Then
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 0, Pos_Yo = 0 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 0 "
     'Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 0, Pos_Yo = 0 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 1 "
     'Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Formato_Propio " _
          & "SET Pos_Xo = 0, Pos_Yo = 0 " _
          & "WHERE TP = 'IF' " _
          & "AND Num = 50 "
     'Ejecutar_SQL_SP sSQL
     SQL2 = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & FA.TC & "' " _
          & "AND Serie = '" & FA.Serie & "' " _
          & "AND Factura BETWEEN " & FA.Desde & " and " & FA.Hasta & " " _
          & "ORDER BY Factura,Cod_CxC "
     Select_Adodc AdoFactura, SQL2
     If AdoFactura.Recordset.RecordCount > 0 Then
       'MsgBox AdoFactura.Recordset.Fields("Cod_CxC")
        FA.Cod_CxC = AdoFactura.Recordset.fields("Cod_CxC")
       'Lineas_De_CxC FA
        FA.TC = DCTipo
        FA.Serie = DCSerie
        FA.Autorizacion = TxtAutorizacion
        Bandera = False
        Evaluar = True
       'MsgBox "..."
        If FA.Desde <= FA.Hasta Then
           If CheqSoloCopia.value = 1 Then
              Imprimir_Facturas_Copias_CxC ListFact, AdoFactura, AdoDetalle, Factura_Desde, Factura_Hasta, FA, True, CBool(CheqMatricula.value), CBool(OpcAsc.value)
           Else
              Imprimir_Facturas_CxC ListFact, FA, True, CBool(CheqMatricula.value), True, CBool(OpcAsc.value)
           End If
        End If
        Facturas_Impresas FA
     End If
  Else
    MsgBox "No se puede imprimir el rando de Facturas"
  End If
End Sub

Public Sub Volver_Autorizar_NC(Optional VolverAutorizar As Boolean)
Dim VuelveAutorizar As Boolean
  If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
     VuelveAutorizar = False
     If Len(FA.Autorizacion_NC) = 13 Then VuelveAutorizar = True
     If Len(FA.Autorizacion_NC) > 13 And VolverAutorizar Then VuelveAutorizar = True
     
    If FA.Factura > 0 And VuelveAutorizar Then
       Mensajes = "Esta seguro que desea volver autorizar la" & vbCrLf & vbCrLf _
                & "Nota de Crédito No. " & FA.Serie_NC & "-" & Format(FA.Nota_Credito, String(9, "0")) & vbCrLf & vbCrLf _
                & "de la Factura No. " & FA.Serie & "-" & Format(FA.Factura, String(9, "0"))
       Titulo = "AUTORIZACION NOTA DE CREDITO"
       If BoxMensaje = vbYes Then
          RatonReloj
          FA.ClaveAcceso_NC = Ninguno
          FA.Estado_SRI_NC = Ninguno
          FA.Hora_NC = Ninguno
          sSQL = "UPDATE Trans_Abonos " _
               & "SET Fecha_Aut_NC = #" & BuscarFecha(FA.Fecha_Aut_NC) & "#," _
               & "Serie_NC = '" & FA.Serie_NC & "'," _
               & "Autorizacion_NC = '" & FA.Autorizacion_NC & "'," _
               & "Secuencial_NC = " & FA.Nota_Credito & "," _
               & "Clave_Acceso_NC = '" & FA.ClaveAcceso_NC & "'," _
               & "Estado_SRI_NC = '" & FA.Estado_SRI_NC & "'," _
               & "Hora_Aut_NC = '" & FA.Hora_NC & "' " _
               & "WHERE Factura = " & FA.Factura & " " _
               & "AND TP = '" & FA.TC & "' " _
               & "AND Serie = '" & FA.Serie & "' " _
               & "AND Autorizacion = '" & FA.Autorizacion & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Ejecutar_SQL_SP sSQL

          SRI_Crear_Clave_Acceso_Nota_Credito FA, True, CBool(CheqClaveAcceso.value)
          TxtXML = "(" & SRI_Autorizacion.Estado_SRI & ") " & SRI_Autorizacion.Error_SRI
          
          'MsgBox FA.Autorizacion_NC
          If Len(FA.Autorizacion_NC) > 13 Then
             sSQL = "UPDATE Trans_Abonos " _
                  & "SET Fecha_Aut_NC = #" & BuscarFecha(FA.Fecha_Aut_NC) & "#, " _
                  & "Autorizacion_NC = '" & FA.Autorizacion_NC & "', " _
                  & "Clave_Acceso_NC = '" & FA.ClaveAcceso_NC & "', " _
                  & "Estado_SRI_NC = '" & FA.Estado_SRI_NC & "', " _
                  & "Hora_Aut_NC = '" & FA.Hora_NC & "' " _
                  & "WHERE Factura = " & FA.Factura & " " _
                  & "AND TP = '" & FA.TC & "' " _
                  & "AND Serie = '" & FA.Serie & "' " _
                  & "AND Autorizacion = '" & FA.Autorizacion & "' " _
                  & "AND Serie_NC = '" & FA.Serie_NC & "' " _
                  & "AND Secuencial_NC = '" & FA.Nota_Credito & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' "
            'MsgBox sSQL
             Ejecutar_SQL_SP sSQL
          End If
          RatonNormal
       End If
    Else
       MsgBox "Este Tipo de Nota de Credito no es electronica"
    End If
  Else
    RatonNormal
    MsgBox MensajeNoAutorizarCE
  End If
End Sub

Public Sub Volver_Autorizar_GR()
  If CFechaLong(FechaSistema) <= CFechaLong(Fecha_CE) Then
    If FA.Factura > 0 And Len(FA.Autorizacion_GR) >= 13 Then
       Mensajes = "Esta seguro que desea volver autorizar la" & vbCrLf & vbCrLf _
                & "Guia de Remision No. " & FA.Serie_GR & "-" & Format(FA.Remision, String(9, "0")) & vbCrLf & vbCrLf _
                & "de la Factura No. " & FA.Serie & "-" & Format(FA.Factura, String(9, "0"))
       Titulo = "AUTORIZACION GUIA DE REMISION"
       If BoxMensaje = vbYes Then
          RatonReloj
          SRI_Crear_Clave_Acceso_Guia_Remision FA, True, CBool(CheqClaveAcceso.value)
          TxtXML = "(" & SRI_Autorizacion.Estado_SRI & ") " & SRI_Autorizacion.Error_SRI
          sSQL = "UPDATE Facturas_Auxiliares " _
               & "SET Fecha_Aut_GR = #" & BuscarFecha(FA.Fecha_Aut_GR) & "#," _
               & "Autorizacion_GR = '" & FA.Autorizacion_GR & "'," _
               & "Clave_Acceso_GR = '" & FA.ClaveAcceso_GR & "'," _
               & "Estado_SRI_GR = '" & FA.Estado_SRI_GR & "'," _
               & "Hora_Aut_GR = '" & FA.Hora_GR & "' " _
               & "WHERE Factura = " & FA.Factura & " " _
               & "AND TC = '" & FA.TC & "' " _
               & "AND Serie = '" & FA.Serie & "' " _
               & "AND Autorizacion = '" & FA.Autorizacion & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Ejecutar_SQL_SP sSQL
          RatonNormal
       End If
    Else
       MsgBox "Este Tipo de Guia de Remision no es electronica"
    End If
  Else
    RatonNormal
    MsgBox MensajeNoAutorizarCE
  End If
End Sub

Public Sub Actualizar_Representantes()
  If ClaveAdministrador Then
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     Control_Procesos "F", "Actualiza Representantes desde " & Factura_Desde & " a la " & Factura_Hasta
     If (Factura_Hasta - Factura_Desde) >= 0 Then
        RatonReloj
        If SQL_Server Then
           SQL2 = "UPDATE Facturas " _
                & "SET RUC_CI=CM.Cedula_R, Razon_Social=CM.Representante, TB=CM.TD " _
                & "FROM Facturas As F,Clientes_Matriculas AS CM "
        Else
           SQL2 = "UPDATE Facturas As F,Clientes_Matriculas AS CM " _
                & "SET F.RUC_CI=CM.Cedula_R, F.Razon_Social=CM.Representante, F.TB=CM.TD "
        End If
        SQL2 = SQL2 _
             & "WHERE F.Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " " _
             & "AND F.Item = '" & NumEmpresa & "' " _
             & "AND F.Periodo = '" & Periodo_Contable & "' " _
             & "AND F.Razon_Social = 'CONSUMIDOR FINAL' " _
             & "AND CM.TD IN ('C','R','P') " _
             & "AND F.TC = '" & FA.TC & "' " _
             & "AND F.Serie = '" & FA.Serie & "' " _
             & "AND LEN(F.Autorizacion) <= 13 " _
             & "AND F.Item = CM.Item " _
             & "AND F.Periodo = CM.Periodo " _
             & "AND F.CodigoC = CM.Codigo "
        Ejecutar_SQL_SP SQL2
        RatonNormal
        MsgBox "Proceso Terminado"
     Else
        MsgBox "No se puede procesar el rando de Facturas"
     End If
  End If
End Sub

Public Sub Actualizar_Kardex()
Dim AdoDBReceta As ADODB.Recordset
  If ClaveAdministrador Then
     Mensajes = "Esta Seguro que desea re-activar kardex del " & vbCrLf _
              & "Documento No. " & FA.TC & ": " & FA.Serie & "-" & FA.Factura
     Titulo = " FORMULARIO DE RE-ACTIVACION"
     If BoxMensaje = vbYes Then
        With AdoDetalle.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
             sSQL = "DELETE * " _
                  & "FROM Trans_Kardex " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND TC = '" & FA.TC & "' " _
                  & "AND Serie = '" & FA.Serie & "' " _
                  & "AND Factura = " & FA.Factura & " " _
                  & "AND SUBSTRING(Detalle, 1, 3) = 'FA:' "
             Ejecutar_SQL_SP sSQL
             Do While Not .EOF
                If Leer_Codigo_Inv(.fields("Codigo"), FA.Fecha, .fields("CodBodega"), .fields("CodMarca")) Then
                   If DatInv.Costo > 0 Then
                      SetAdoAddNew "Trans_Kardex"
                      SetAdoFields "T", Normal
                      SetAdoFields "TC", FA.TC
                      SetAdoFields "Serie", FA.Serie
                      SetAdoFields "Fecha", FA.Fecha
                      SetAdoFields "Factura", FA.Factura
                      SetAdoFields "Codigo_P", FA.CodigoC
                      SetAdoFields "CodBodega", .fields("CodBodega")
                      SetAdoFields "CodMarca", .fields("CodMarca")
                      SetAdoFields "Codigo_Inv", .fields("Codigo")
                      SetAdoFields "CodigoL", FA.Cod_CxC
                      SetAdoFields "Lote_No", .fields("Lote_No")
                      SetAdoFields "Fecha_Fab", DatInv.Fecha_Fab
                      SetAdoFields "Fecha_Exp", DatInv.Fecha_Exp
                      SetAdoFields "Procedencia", .fields("Procedencia")
                      SetAdoFields "Modelo", .fields("Modelo")
                      SetAdoFields "Serie_No", .fields("Serie_No")
                      SetAdoFields "Total_IVA", .fields("Total_IVA")
                      SetAdoFields "Porc_C", FA.Porc_C
                      SetAdoFields "Salida", .fields("Cantidad")
                      SetAdoFields "PVP", .fields("Precio")
                      SetAdoFields "Valor_Unitario", .fields("Precio")
                      SetAdoFields "Costo", DatInv.Costo
                      SetAdoFields "Valor_Total", Redondear(.fields("Cantidad") * .fields("Precio"), 2)
                      SetAdoFields "Total", Redondear(.fields("Cantidad") * DatInv.Costo, 2)
                      SetAdoFields "Detalle", MidStrg("FA: " & FA.Cliente, 1, 100)
                      SetAdoFields "Codigo_Barra", DatInv.Codigo_Barra
                     'SetAdoFields "Orden_No", .Fields("Numero")
                      SetAdoFields "Cta_Inv", DatInv.Cta_Inventario
                      SetAdoFields "Contra_Cta", DatInv.Cta_Costo_Venta
                      SetAdoFields "Item", NumEmpresa
                      SetAdoFields "Periodo", Periodo_Contable
                      SetAdoFields "CodigoU", CodigoUsuario
                      SetAdoUpdate
                   End If
                   
                End If
               'Salida si es por recetas
                sSQL = "SELECT Codigo_Receta, Cantidad, Costo, ID " _
                     & "FROM Catalogo_Recetas " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo_PP = '" & .fields("Codigo") & "' " _
                     & "AND TC = 'P' " _
                     & "ORDER BY Codigo_Receta "
                Select_AdoDB AdoDBReceta, sSQL
                If AdoDBReceta.RecordCount > 0 Then
                   Do While Not AdoDBReceta.EOF
                      If Leer_Codigo_Inv(AdoDBReceta.fields("Codigo_Receta"), FechaSistema, .fields("CodBodega"), .fields("CodMarca")) Then
                        'MsgBox .fields("Codigo") & vbCrLf & AdoDBReceta.fields("Codigo_Receta") & vbCrLf & DatInv.Costo
                         If DatInv.Costo > 0 Then
                            CantidadAnt = .fields("Cantidad") * AdoDBReceta.fields("Cantidad")
                            ValorTotal = Redondear(CantidadAnt * DatInv.Costo, 2)
                            SetAdoAddNew "Trans_Kardex"
                            SetAdoFields "T", Normal
                            SetAdoFields "TC", FA.TC
                            SetAdoFields "Serie", FA.Serie
                            SetAdoFields "Fecha", FA.Fecha
                            SetAdoFields "Factura", FA.Factura
                            SetAdoFields "Codigo_P", FA.CodigoC
                            SetAdoFields "CodBodega", .fields("CodBodega")
                            SetAdoFields "CodMarca", .fields("CodMarca")
                            SetAdoFields "Codigo_Inv", AdoDBReceta.fields("Codigo_Receta")
'''                                  SetAdoFields "CodigoL", FA.Cod_CxC
'''                                  SetAdoFields "Lote_No", .fields("Lote_No")
'''                                  SetAdoFields "Fecha_Fab", .fields("Fecha_Fab")
'''                                  SetAdoFields "Fecha_Exp", .fields("Fecha_Exp")
'''                                  SetAdoFields "Procedencia", .fields("Procedencia")
'''                                  SetAdoFields "Modelo", .fields("Modelo")
'''                                  SetAdoFields "Serie_No", .fields("Serie_No")
'''                                  SetAdoFields "Porc_C", .fields("Porc_C")
                            SetAdoFields "PVP", DatInv.Costo
                            SetAdoFields "Valor_Unitario", DatInv.Costo
                            SetAdoFields "Salida", CantidadAnt
                            SetAdoFields "Valor_Total", ValorTotal
                            SetAdoFields "Costo", DatInv.Costo
                            SetAdoFields "Total", ValorTotal
                            SetAdoFields "Detalle", MidStrg("FA: RE-" & FA.Cliente, 1, 100)
                            SetAdoFields "Cta_Inv", DatInv.Cta_Inventario
                            SetAdoFields "Contra_Cta", DatInv.Cta_Costo_Venta
                            SetAdoFields "Item", NumEmpresa
                            SetAdoFields "Periodo", Periodo_Contable
                            SetAdoFields "CodigoU", CodigoUsuario
                            SetAdoUpdate
                         End If
                      End If
                      AdoDBReceta.MoveNext
                   Loop
                End If
                AdoDBReceta.Close
                
               .MoveNext
             Loop
         End If
        End With
        MsgBox "Proceso Terminado"
     End If
  End If
End Sub

Public Sub Liberar_FA_SRI()
Dim IdFact As Long
  If ClaveAdministrador Then
     RatonReloj
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     If (Factura_Hasta - Factura_Desde) >= 0 Then
        Mensajes = "Esta seguro que desea liberar en masa " & FA.TC & "-" & FA.Serie & vbCrLf _
                 & "desde " & Factura_Desde & " a la " & Factura_Hasta & vbCrLf
        Titulo = "LIBERACION DE FACTURAS SRI"
        If BoxMensaje = vbYes Then
           Control_Procesos "F", "Liberacion de Facturas en masa desde " & Factura_Desde & " a la " & Factura_Hasta
           RatonReloj
           sSQL = "UPDATE Facturas " _
                & "SET Autorizacion = '" & RUC & "', Clave_Acceso = '.', Estado_SRI = '.' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND LEN(Autorizacion) >= 13 " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "UPDATE Detalle_Factura " _
                & "SET Autorizacion = '" & RUC & "' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND LEN(Autorizacion) >= 13 " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "UPDATE Trans_Abonos " _
                & "SET Autorizacion = '" & RUC & "' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND LEN(Autorizacion) >= 13 " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
           RatonNormal
           MsgBox "Proceso Terminado con exito, se volvera a intentar autorizar los documentos"
           Volver_Autorizar_Pendientes
           RatonNormal
           MsgBox "Si despues de este proceso no se autorizaron los documentos, ejecute el modulos de seteos, " _
                & "e intente actualizar los documentos, o llame al Servicio al Cliente de DiskCover Sistem"
        End If
     End If
  End If
End Sub

Public Sub Anular_en_masa()
Dim Motivo As String
  If ClaveAdministrador Then
     RatonNormal
     If Val(TextFDesde) <= 0 Then TextFDesde = "0"
     If Val(TextFHasta) <= 0 Then TextFHasta = "0"
     TextoValido TextFDesde
     TextoValido TextFHasta
     Factura_Desde = Val(TextFDesde)
     Factura_Hasta = Val(TextFHasta)
     Progreso_Barra.Mensaje_Box = "Anulacion de Facturas en masa"
     Progreso_Iniciar
     If (Factura_Hasta - Factura_Desde) >= 0 Then
        Mensajes = "Esta Seguro de Anular las " & FA.TC & " de la Serie: " & FA.Serie & vbCrLf & vbCrLf _
                 & "desde la " & Factura_Desde & " hasta la " & Factura_Hasta & "en bloque " & vbCrLf & vbCrLf _
                 & "DIGITE EL MOTIVO DE LA ANULACION:"
        Titulo = "Formulario de Anulación"
        Motivo = InputBox(Mensajes, UCase(Titulo), "Anulacion en masa")
        If Len(Motivo) > 1 Then
           Control_Procesos "F", "Anulacion de Facturas desde la " & Factura_Desde & " hasta la " & Factura_Hasta
           RatonReloj
           sSQL = "UPDATE Facturas " _
                & "SET T = 'A', Nota = 'Motivo de la Anulacion: " & Motivo & ".' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "UPDATE Detalle_Factura " _
                & "SET T = 'A' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TC = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
     
           sSQL = "DELETE * " _
                & "FROM Trans_Abonos " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND TP = '" & FA.TC & "' " _
                & "AND Serie = '" & FA.Serie & "' " _
                & "AND Factura BETWEEN " & Factura_Desde & " AND " & Factura_Hasta & " "
           Ejecutar_SQL_SP sSQL
           RatonNormal
           MsgBox "Proceso realizado con exito"
        Else
           MsgBox "Proceso de anulacion cancelado"
        End If
     End If
     Progreso_Final
  End If
End Sub

Public Sub Actualizar_NC_Kardex()
Dim AdoAuxDB As ADODB.Recordset
    With AdoDetalle.Recordset
     If .RecordCount > 0 Then
         FA.Serie_NC = .fields("Serie_NC")
         FA.Nota_Credito = .fields("Secuencial_NC")
         FA.Autorizacion_NC = .fields("Autorizacion_NC")
         CodigoB = .fields("Banco")
         CodigoA = .fields("Cheque")
         If .fields("Banco") = "NOTA DE CREDITO" Then
             Titulo = "CONFIRMACION DE ELIMINACION"
             Mensajes = "¿Realmente desea Reactivar la Nota de Credito en el Kardex:" & vbCrLf _
                      & "Fecha: " & .fields("Fecha") & vbCrLf _
                      & "Factura No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") & vbCrLf _
                      & " " _
                      & " "
             If Len(FA.Serie_NC) = 6 And FA.Nota_Credito > 0 And Len(FA.Autorizacion_NC) >= 8 Then
                Mensajes = Mensajes & "y ademas contiene la siguiente: " & vbCrLf _
                         & "NOTA DE CREDITO: " & FA.Serie_NC & "-" & Format(FA.Nota_Credito, "000000000") & vbCrLf
             End If
             If BoxMensaje = vbYes Then
                sSQL = "DELETE * " _
                     & "FROM Trans_Kardex " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND TC = '" & FA.TC & "' " _
                     & "AND Serie = '" & FA.Serie & "' " _
                     & "AND Factura = " & FA.Factura & " " _
                     & "AND SUBSTRING(Detalle,1,3) = 'NC:' "
                Ejecutar_SQL_SP sSQL
                
                sSQL = "SELECT " & Full_Fields("Detalle_Nota_Credito") & "  " _
                     & "FROM Detalle_Nota_Credito " _
                     & "WHERE TC = '" & FA.TC & "' " _
                     & "AND Factura = " & FA.Factura & " " _
                     & "AND Serie_FA = '" & FA.Serie & "' " _
                     & "AND Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Serie = '" & FA.Serie_NC & "' " _
                     & "AND Secuencial = " & FA.Nota_Credito & " "
                Select_AdoDB AdoAuxDB, sSQL
                If AdoAuxDB.RecordCount > 0 Then
                    Do While Not AdoAuxDB.EOF
                       Stock_Actual_Inventario FA.Fecha_NC, AdoAuxDB.fields("Codigo_Inv"), AdoAuxDB.fields("CodBodega")
                       ValorTotal = Redondear(ValorUnit * AdoAuxDB.fields("Cantidad"), 2)
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
                       SetAdoFields "Codigo_Barra", Cod_Barra
                       SetAdoFields "CodBodega", AdoAuxDB.fields("CodBodega")
                       'SetAdoFields "CodMarca", AdoAuxDB.fields("CodMar")
                       SetAdoFields "Codigo_Inv", AdoAuxDB.fields("Codigo_Inv")
                       SetAdoFields "Total_IVA", AdoAuxDB.fields("Total_IVA")
                       SetAdoFields "Entrada", AdoAuxDB.fields("Cantidad")
                       SetAdoFields "PVP", AdoAuxDB.fields("Precio") 'SubTotalCosto
                       SetAdoFields "Valor_Unitario", ValorUnit 'SubTotalCosto
                       SetAdoFields "Costo", ValorUnit
                       SetAdoFields "Valor_Total", ValorTotal
                       SetAdoFields "Total", ValorTotal
                       SetAdoFields "Descuento", AdoAuxDB.fields("Descuento")
                       SetAdoFields "Detalle", "NC: " + FA.Serie_NC + "-" + Format(FA.Nota_Credito, "000000000") + " -" + MidStrg(FA.Cliente, 1, 79)
                       SetAdoFields "Cta_Inv", Cta_Aux_Inv
                       SetAdoFields "Contra_Cta", AdoAuxDB.fields("Cta_Devolucion")
                       SetAdoFields "Item", NumEmpresa
                       SetAdoFields "Periodo", Periodo_Contable
                       SetAdoFields "CodigoU", CodigoUsuario
                       SetAdoUpdate
                       AdoAuxDB.MoveNext
                    Loop
                End If
                AdoAuxDB.Close
             End If
         Else
             MsgBox "No existe Nota de Credito que actualizar al Kardex"
         End If
     Else
         MsgBox "No existe datos que procesar"
     End If
   End With
End Sub

''''    SubSQL = "SELECT SUM(Abono) " _
''''           & "FROM Trans_Abonos " _
''''           & "WHERE Trans_Abonos.TP = Facturas.TC " _
''''           & "AND Trans_Abonos.Item = Facturas.Item " _
''''           & "AND Trans_Abonos.Periodo = Facturas.Periodo " _
''''           & "AND Trans_Abonos.Factura = Facturas.Factura " _
''''           & "AND Trans_Abonos.CodigoC = Facturas.CodigoC " _
''''           & "AND Trans_Abonos.Serie = Facturas.Serie " _
''''           & "AND Trans_Abonos.Autorizacion = Facturas.Autorizacion "
''''
''''    For IDMes = 1 To 12
''''        sSQL = "UPDATE Facturas " _
''''             & "SET Abonos_MN = (" & SubSQL & ") " _
''''             & "WHERE Item = '" & NumEmpresa & "' " _
''''             & "AND Periodo = '" & Periodo_Contable & "' " _
''''             & "AND TC = '" & FA.TC & "' " _
''''             & "AND Serie = '" & FA.Serie & "' " _
''''             & "AND MONTH(Fecha) = " & IDMes & " " _
''''             & "AND T <> 'A' "
''''        Ejecutar_SQL_SP sSQL
''''    Next IDMes
''''
''''    sSQL = "UPDATE Facturas " _
''''         & "SET Abonos_MN = 0 " _
''''         & "WHERE Item = '" & NumEmpresa & "' " _
''''         & "AND Periodo = '" & Periodo_Contable & "' " _
''''         & "AND TC = '" & FA.TC & "' " _
''''         & "AND Serie = '" & FA.Serie & "' " _
''''         & "AND Abonos_MN IS NULL "
''''    Ejecutar_SQL_SP sSQL

