VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTAS"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "FVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10620
   Begin VB.ComboBox CAño 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FVentas.frx":0696
      Left            =   210
      List            =   "FVentas.frx":0698
      TabIndex        =   60
      Text            =   "2000"
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox CModificacion 
      DataSource      =   "AdoTransVentas"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8070
   End
   Begin VB.ComboBox CMes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   12
      Visible         =   0   'False
      Width           =   1416
   End
   Begin VB.Frame FrmRetencion 
      Caption         =   "RETENCIONES DE IVA POR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   984
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   7350
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FVentas.frx":069A
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   525
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FVentas.frx":06B7
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox ChRetS 
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   630
         Width           =   1170
      End
      Begin VB.CheckBox ChRetB 
         Caption         =   "Bienes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame FrmTipoComprob 
      Height          =   984
      Left            =   7560
      TabIndex        =   7
      Top             =   324
      Width           =   1956
      Begin VB.TextBox TxtNumeroC 
         Alignment       =   1  'Right Justify
         Height          =   336
         Left            =   735
         MaxLength       =   9
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "FVentas.frx":06D7
         ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
         Top             =   525
         Width           =   1080
      End
      Begin VB.ComboBox CTP 
         Height          =   315
         Left            =   105
         TabIndex        =   8
         ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
         Top             =   525
         Width           =   660
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TP      NUMERO"
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
         TabIndex        =   79
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   9555
      Picture         =   "FVentas.frx":06DB
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Grabar"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   9555
      Picture         =   "FVentas.frx":09E5
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Salir"
      Top             =   1155
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCClienteV 
      Bindings        =   "FVentas.frx":0E27
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   210
      TabIndex        =   10
      ToolTipText     =   "En este combo de selección se despliega una lista de todos los Clientes"
      Top             =   1575
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   556
      _Version        =   393216
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
   Begin TabDlg.SSTab SSTVentas 
      Height          =   4455
      Left            =   105
      TabIndex        =   14
      Top             =   1995
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- COMPROBANTE DE VENTA"
      TabPicture(0)   =   "FVentas.frx":0E41
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrmIva"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&2.- INSERTAR CONCEPTO AIR"
      TabPicture(1)   =   "FVentas.frx":0E5D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "INGRESE LOS DATOS DE LA FACTURA, NOTA DE VENTA, ETC. ______________________ FORMULARIO 104"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1800
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Width           =   9990
         Begin VB.CommandButton CmdAir 
            Caption         =   "&AIR"
            Height          =   444
            Left            =   9240
            Picture         =   "FVentas.frx":0E79
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Se ubica en la pestaña de Retenciones"
            Top             =   315
            Width           =   552
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   19
            Text            =   "0000001"
            ToolTipText     =   $"FVentas.frx":139F
            Top             =   1365
            Width           =   1065
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   17
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   1365
            Width           =   540
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   1365
            Width           =   645
         End
         Begin VB.TextBox TxtBaseImpV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   5145
            TabIndex        =   22
            Text            =   "0.00"
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1485
         End
         Begin VB.TextBox TxtNumComprobante 
            Height          =   336
            Left            =   7560
            MaxLength       =   7
            TabIndex        =   16
            Top             =   525
            Width           =   1275
         End
         Begin VB.TextBox TxtBaseImpGravV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   23
            Text            =   "0.00"
            ToolTipText     =   $"FVentas.frx":1442
            Top             =   1365
            Width           =   1590
         End
         Begin VB.TextBox TxtBaseImpoIceV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8295
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "FVentas.frx":14EA
            ToolTipText     =   $"FVentas.frx":14EF
            Top             =   1365
            Width           =   1590
         End
         Begin MSDataListLib.DataCombo DCTipoComprobanteV 
            Bindings        =   "FVentas.frx":1581
            DataSource      =   "AdoTipoComprobante"
            Height          =   315
            Left            =   105
            TabIndex        =   15
            ToolTipText     =   $"FVentas.frx":15A2
            Top             =   525
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiV 
            Height          =   330
            Left            =   2415
            TabIndex        =   20
            ToolTipText     =   $"FVentas.frx":164A
            Top             =   1365
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBFechaRegistroV 
            Height          =   330
            Left            =   3780
            TabIndex        =   21
            ToolTipText     =   $"FVentas.frx":16F6
            Top             =   1365
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO DE COMPROBANTE"
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
            TabIndex        =   73
            Top             =   315
            Width           =   7365
         End
         Begin VB.Label Label7 
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
            Height          =   225
            Left            =   7560
            TabIndex        =   72
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TARIFA 0%"
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
            Left            =   5145
            TabIndex        =   71
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TARIFA 12%"
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
            TabIndex        =   70
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR ICE"
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
            TabIndex        =   69
            Top             =   1050
            Width           =   1590
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SERIE Y COMPROBAN"
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
            TabIndex        =   68
            Top             =   1050
            Width           =   2220
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMISION"
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
            TabIndex        =   67
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CADUCIDAD"
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
            Left            =   3780
            TabIndex        =   66
            Top             =   1050
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Retención Presuntiva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2835
         TabIndex        =   28
         Top             =   2205
         Width           =   2535
         Begin VB.OptionButton OpcRetNo 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1470
            TabIndex        =   30
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton OpcRetSi 
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   29
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Frame FrmIva 
         Caption         =   "I.V.A. Presuntivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   25
         Top             =   2205
         Width           =   2535
         Begin VB.OptionButton OpcIvaNo 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1470
            TabIndex        =   27
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton OpcIvaSi 
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   26
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "RETENCION DEL IVA POR BIENES Y/O SERVICIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   5460
         TabIndex        =   36
         Top             =   2205
         Width           =   4752
         Begin VB.TextBox TxtIvaBienMonIvaV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   37
            Text            =   "FVentas.frx":177E
            ToolTipText     =   $"FVentas.frx":1783
            Top             =   630
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaBienValRetV 
            Enabled         =   0   'False
            Height          =   336
            Left            =   1680
            TabIndex        =   39
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaSerMonIvaV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3156
            MultiLine       =   -1  'True
            TabIndex        =   40
            Text            =   "FVentas.frx":1822
            ToolTipText     =   $"FVentas.frx":1827
            Top             =   630
            Width           =   1380
         End
         Begin VB.TextBox TxtIvaSerValRetV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3156
            TabIndex        =   42
            Text            =   " "
            Top             =   1470
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBienV 
            Bindings        =   "FVentas.frx":18BD
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1680
            TabIndex        =   38
            ToolTipText     =   $"FVentas.frx":18DB
            Top             =   1050
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServV 
            Bindings        =   "FVentas.frx":1967
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3150
            TabIndex        =   41
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   1050
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " BIENES"
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
            Left            =   1680
            TabIndex        =   65
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label20 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SERVICIOS"
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
            Left            =   3150
            TabIndex        =   64
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label21 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
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
            Left            =   315
            TabIndex        =   63
            Top             =   630
            Width           =   1380
         End
         Begin VB.Label Label22 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PORCENTAJE"
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
            Left            =   315
            TabIndex        =   62
            Top             =   1050
            Width           =   1380
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR RET."
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
            Left            =   315
            TabIndex        =   61
            Top             =   1470
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "PORCENTAJE DE LAS BASES IMPONIBLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1230
         Left            =   210
         TabIndex        =   31
         Top             =   3045
         Width           =   5145
         Begin VB.TextBox TxtMontoIvaV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   33
            ToolTipText     =   "Este valor se calcula automaticamente, es el resultado de aplicarle un porcentaje IVA a la Base Imponible gravada"
            Top             =   315
            Width           =   1485
         End
         Begin VB.TextBox TxtMontoIceV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3465
            TabIndex        =   35
            Top             =   735
            Width           =   1485
         End
         Begin MSDataListLib.DataCombo DCPorcenIvaV 
            Bindings        =   "FVentas.frx":1988
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   32
            ToolTipText     =   $"FVentas.frx":19A0
            Top             =   315
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIceV 
            Bindings        =   "FVentas.frx":1A32
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   34
            ToolTipText     =   $"FVentas.frx":1A4A
            Top             =   735
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " I.V.A."
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
            TabIndex        =   77
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ICE"
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
            TabIndex        =   76
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR I.V.A."
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
            Left            =   1995
            TabIndex        =   75
            Top             =   315
            Width           =   1485
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR ICE"
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
            Left            =   1995
            TabIndex        =   74
            Top             =   735
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "INGRESE LOS DATOS DE LA RETENCION _________________________________________ FORMULARIO 103"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4020
         Left            =   -74895
         TabIndex        =   44
         Top             =   315
         Width           =   10155
         Begin VB.TextBox TxtValConAV 
            Enabled         =   0   'False
            Height          =   330
            Left            =   8715
            TabIndex        =   55
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConAV 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8085
            TabIndex        =   54
            Top             =   1470
            Width           =   645
         End
         Begin VB.TextBox TxtBimpConAV 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   53
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtSumatoria 
            Alignment       =   1  'Right Justify
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
            Height          =   336
            Left            =   8820
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   840
            Width           =   1170
         End
         Begin VB.TextBox TxtNumTresComRetV 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   49
            Text            =   "0000001"
            ToolTipText     =   $"FVentas.frx":1ADB
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox TxtNumDosComRetV 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   48
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRetV 
            Height          =   336
            Left            =   108
            MaxLength       =   3
            TabIndex        =   47
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtTotalReten 
            Alignment       =   1  'Right Justify
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
            Height          =   336
            Left            =   8715
            TabIndex        =   57
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   3570
            Width           =   1275
         End
         Begin VB.TextBox TxtNumUnoAutComRetV 
            Height          =   330
            Left            =   2415
            MaxLength       =   37
            TabIndex        =   50
            ToolTipText     =   $"FVentas.frx":1B7D
            Top             =   840
            Width           =   1590
         End
         Begin VB.CheckBox ChRetF 
            Caption         =   "Retención en la Fuente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   45
            Top             =   315
            Width           =   2328
         End
         Begin MSDataListLib.DataCombo DCConceptoRetV 
            Bindings        =   "FVentas.frx":1C09
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   52
            Top             =   1470
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataGridLib.DataGrid DGConceptoAirV 
            Bindings        =   "FVentas.frx":1C26
            Height          =   1545
            Left            =   105
            TabIndex        =   56
            Top             =   1890
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   2725
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Datos Ingresados"
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
                  LCID            =   3082
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
                  LCID            =   3082
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
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FVentas.frx":1C42
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2520
            TabIndex        =   46
            Top             =   315
            Visible         =   0   'False
            Width           =   7470
            _ExtentX        =   13176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VALOR RET."
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
            Left            =   8715
            TabIndex        =   88
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PORC"
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
            Left            =   8085
            TabIndex        =   87
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CODIGO DE RETENCION"
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
            TabIndex        =   86
            Top             =   1260
            Width           =   6630
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SUMATORIA"
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
            TabIndex        =   85
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label30 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Autorización"
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
            Left            =   2415
            TabIndex        =   84
            Top             =   630
            Width           =   1590
         End
         Begin VB.Label Label29 
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
            Height          =   225
            Left            =   1260
            TabIndex        =   83
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie"
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
            TabIndex        =   82
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Retenciones"
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
            TabIndex        =   81
            Top             =   3570
            Width           =   2010
         End
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " BASE IMP."
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
            Left            =   6720
            TabIndex        =   80
            Top             =   1260
            Width           =   1380
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoTransVentas 
      Height          =   330
      Left            =   210
      Top             =   5190
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "TransCompras"
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
      Left            =   210
      Top             =   4245
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "RetencionFuente"
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
   Begin MSAdodcLib.Adodc AdoAsientoVentas 
      Height          =   330
      Left            =   210
      Top             =   4560
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AsientoCompras"
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienesCC 
      Height          =   330
      Left            =   210
      Top             =   5820
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "RetencionIvaBienes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   210
      Top             =   5505
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "RetencionIvaServicios"
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
      Left            =   3045
      Top             =   2040
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   3045
      Top             =   2670
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "AsientoAirVentas"
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
   Begin MSAdodcLib.Adodc AdoConceptoRet 
      Height          =   330
      Left            =   210
      Top             =   3930
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoConceptoAir"
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
   Begin MSAdodcLib.Adodc AdoPorIce 
      Height          =   330
      Left            =   210
      Top             =   3300
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoPorIce"
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
   Begin MSAdodcLib.Adodc AdoPorIva 
      Height          =   330
      Left            =   210
      Top             =   3615
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoPorIva"
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   210
      Top             =   2655
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoRetenBienes"
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   2325
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoTipoComp"
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
   Begin MSAdodcLib.Adodc AdoTipoIdentificacion 
      Height          =   330
      Left            =   210
      Top             =   1995
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoTipoIden"
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
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   210
      Top             =   4875
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoSustento"
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
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   3045
      Top             =   2355
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "TransAirventas"
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicios 
      Height          =   330
      Left            =   210
      Top             =   2970
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
      Caption         =   "AdoRetenServicios"
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
      Left            =   3045
      Top             =   3045
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "AdoAux"
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
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR / BENEFICIARIO"
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
      TabIndex        =   78
      Top             =   1365
      Width           =   9255
   End
   Begin VB.Label LblNumIdentV 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   330
      Left            =   7560
      TabIndex        =   13
      Top             =   1575
      Width           =   1905
   End
   Begin VB.Label LblTD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   7245
      TabIndex        =   11
      Top             =   1575
      Width           =   330
   End
End
Attribute VB_Name = "FVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OP As Boolean
Dim MBFecha As MaskEdBox
Dim Cap1, Captc, Espder, Espizq, ValorR, ValorP, Ct, CodProv, ch, Ch1 As String
Dim Rb, Rs, Rf, cod As Byte
Dim ac, CalIsMi, CalmIva, CalIbMi As Single

Private Sub ChRetB_Click()
    If ChRetB.value <> 0 Then
       ch = 1
       Ch1 = "B"
       DCRetIBienes.Visible = True
       TxtIvaBienMonIvaV.Enabled = True
       DCPorcenRetenIvaBienV.Enabled = True
       TxtIvaBienValRetV.Enabled = True
    Else
       Ch1 = "S"
       TxtIvaBienMonIvaV.Enabled = False
       DCPorcenRetenIvaBienV.Enabled = False
       TxtIvaBienValRetV.Enabled = False
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub ChRetF_Click()
  If ChRetF.value <> 0 Then
     DCRetFuente.Visible = True
     TxtNumUnoComRetV.Enabled = True
     TxtNumDosComRetV.Enabled = True
     TxtNumTresComRetV.Enabled = True
     TxtNumUnoAutComRetV.Enabled = True
     DCConceptoRetV.Enabled = True
     TxtBimpConAV.Enabled = True
  Else
     DCRetFuente.Visible = False
     TxtNumUnoComRetV.Enabled = False
     TxtNumDosComRetV.Enabled = False
     TxtNumTresComRetV.Enabled = False
     TxtNumUnoAutComRetV.Enabled = False
     DCConceptoRetV.Enabled = False
     TxtBimpConAV.Enabled = False
  End If
End Sub

Private Sub ChRetS_Click()
  If ChRetS.value <> 0 Then
     ch = 1
     Ch1 = "S"
     DCRetISer.Visible = True
     TxtIvaSerMonIvaV.Enabled = True
     DCPorcenRetenIvaServV.Enabled = True
     TxtIvaSerValRetV.Enabled = True
  Else
     Ch1 = "B"
     DCRetISer.Visible = False
     TxtIvaSerMonIvaV.Enabled = False
     DCPorcenRetenIvaServV.Enabled = False
     TxtIvaSerValRetV.Enabled = False
  End If
  If ChRetB.value <> 0 And ChRetB.value <> 0 Then
     Ch1 = "X"
  End If
End Sub

Private Sub CmdAir_Click()
  SSTVentas.Tab = 1
  TxtNumUnoComRetV.SetFocus
End Sub

Private Sub CmdCerrar_Click()
  'Borra Asiento Ventas
  sSQL = "DELETE * " _
       & "FROM Asiento_Ventas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'V' "
  ConectarAdoExecute sSQL
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim ProvAnt As String
  RatonReloj
  ProvAnt = DCClienteV
  'Valido por si acaso exista algun valor con 0
  TextoValido TxtIvaBienMonIvaV, True, , 2
  TextoValido TxtBaseImpV, True, , 2
  TextoValido TxtBaseImpGravV, True, , 2
  TextoValido TxtBaseImpoIceV, True, , 2
  TextoValido TxtMontoIvaV, True, , 2
  TextoValido TxtMontoIceV, True, , 2
  TextoValido TxtIvaBienMonIvaV, True, , 2
  TextoValido TxtIvaBienValRetV, True, , 2
  TextoValido TxtIvaSerMonIvaV, True, , 2
  TextoValido TxtIvaSerValRetV, True, , 2
 'Borra si encuentra 2 o mas transacciones iguales
  Eliminar_Trans_AT "V", CodigoCliente, TxtNumSerieUno, TxtNumSerieDos, TxtNumSerietres, Autorizacion, "", "", CStr(cod), CMes, CAño, Ln_SRI
  Eliminar_Trans_Air "V", CodigoCliente, CMes, CAño, Ln_SRI
 'Pregunto antes de grabar
  Titulo = "Grabar Ventas"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
     'Grabacion de los Datos
     Grabacion
     If Ln_SRI < 0 Then
        Titulo = "Grabar Ventas"
        Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
                 & "Desea ingresar otra transacción"""
        If BoxMensaje = vbYes Then
           SSTVentas.Tab = 0
           Limpiar_Controles
           DCClienteV = ProvAnt
           CTP.SetFocus
        Else
           Unload FVentas
        End If
     Else
        MsgBox "Los Datos fueron grabados correctamente"
        Unload FVentas
     End If
  Else
     ChRetB.SetFocus
  End If
End Sub
                            
Private Sub CMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMes_LostFocus()
  Modificacion_AT "V", CModificacion, CMes, CAño
End Sub

Private Sub CModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CModificacion_LostFocus()
    Carga_ConceptosRetencion MBFechaEmiV
    Limpiar_Controles
    Encerar_Var
    'Cargo los datos para la modificación
    CodigoCliente = Ninguno
    I = Val(SinEspaciosIzq(CModificacion))
    Cadena = SinEspaciosIzq(CModificacion)
    NombreCliente = Trim$(Mid$(CModificacion, Len(Cadena) + 1, Len(CModificacion)))
    With AdoClientes.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & NombreCliente & "' ")
         If Not .EOF Then CodigoCliente = .Fields("Codigo")
     End If
    End With
    sSQL = "SELECT * " _
         & "FROM Trans_Ventas " _
         & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Linea_SRI = " & I & " " _
         & "AND IdProv = '" & CodigoCliente & "' " _
         & "ORDER BY Linea_SRI "
    SelectAdodc AdoTransVentas, sSQL
    With AdoTransVentas.Recordset
     If .RecordCount > 0 Then
         Ln_SRI = I
              'Busco el Proveedor
              CodProv = .Fields("IdProv")
              If AdoClientes.Recordset.RecordCount > 0 Then
                 AdoClientes.Recordset.MoveFirst
                 AdoClientes.Recordset.Find ("Codigo = '" & CodProv & "' ")
                 If Not AdoClientes.Recordset.EOF Then
                    DCClienteV = AdoClientes.Recordset.Fields("Cliente")
                    LblTD = AdoClientes.Recordset.Fields("TD")
                    LblNumIdentV = AdoClientes.Recordset.Fields("CI_RUC")
                 Else
                    MsgBox "Este beneficiario no existe", vbInformation, "Aviso"
                 End If
              End If
              
              If .Fields("IvaPresuntivo") = "S" Then
                  OpcIvaSi.value = True
              Else
                  OpcIvaNo.value = True
              End If
              
              If .Fields("RetPresuntiva") = "S" Then
                  OpcRetSi.value = True
              Else
                  OpcRetNo.value = True
              End If
              
             'Cargo el Comprobante
             sSQL = "SELECT Tipo_Comprobante_Codigo,Descripcion " _
                  & "FROM Tipo_Comprobante " _
                  & "WHERE Tipo_Comprobante_Codigo <> 0 "
             SelectAdodc AdoTipoComprobante, sSQL
             cod = .Fields("TipoComprobante")
             If AdoTipoComprobante.Recordset.RecordCount > 0 Then
                AdoTipoComprobante.Recordset.MoveFirst
                AdoTipoComprobante.Recordset.Find ("Tipo_Comprobante_Codigo = '" & cod & "' ")
                If Not AdoTipoComprobante.Recordset.EOF Then
                   DCTipoComprobanteV = AdoTipoComprobante.Recordset.Fields("Descripcion")
                Else
                   MsgBox "Este Comprobante no existe", vbInformation, "Aviso"
                 End If
             End If
                          
             MBFechaRegistroV = .Fields("FechaRegistro")
             'Es un numero el campo PorcentajeIva y lo convierto a String
             CodPorIva = CStr(.Fields("PorcentajeIva"))
             If AdoPorIva.Recordset.RecordCount > 0 Then
                AdoPorIva.Recordset.MoveFirst
                AdoPorIva.Recordset.Find ("Codigo = '" & CodPorIva & "' ")
                If Not AdoPorIva.Recordset.EOF Then
                   DCPorcenIvaV = AdoPorIva.Recordset.Fields("Porc")
                Else
                   MsgBox "Porcentaje de IVA no existe", vbInformation, "Aviso"
                 End If
             End If
             'Cargo el ICE
              CodPorIce = CStr(.Fields("PorcentajeIce"))
              If AdoPorIce.Recordset.RecordCount > 0 Then
                 AdoPorIce.Recordset.MoveFirst
                 AdoPorIce.Recordset.Find ("Codigo = '" & CodPorIce & "' ")
                 If Not AdoPorIce.Recordset.EOF Then
                    DCPorcenIceV = AdoPorIce.Recordset.Fields("Porc")
                 Else
                    MsgBox "Porcentaje de ICE no existe", vbInformation, "Aviso"
                 End If
              End If
              'Cargo el Codigo retencion Bienes
              CodRetBien = .Fields("PorRetBienes")
              If AdoRetIvaBienes.Recordset.RecordCount > 0 Then
                 AdoRetIvaBienes.Recordset.MoveFirst
                 AdoRetIvaBienes.Recordset.Find ("Codigo = '" & CodRetBien & "' ")
                 If Not AdoRetIvaBienes.Recordset.EOF Then
                    DCPorcenRetenIvaBienV = AdoRetIvaBienes.Recordset.Fields("Porc")
                 Else
                    MsgBox "Código de Retención no existe", vbInformation, "Aviso"
                 End If
              End If
              'Cargo el Codigo retencion Servicios
              CodRetServ = .Fields("PorRetServicios")
              If AdoRetIvaServicios.Recordset.RecordCount > 0 Then
                 AdoRetIvaServicios.Recordset.MoveFirst
                 AdoRetIvaServicios.Recordset.Find ("Codigo = '" & CodRetServ & "' ")
                 If Not AdoRetIvaServicios.Recordset.EOF Then
                    DCPorcenRetenIvaServV = AdoRetIvaServicios.Recordset.Fields("Porc")
                 Else
                    MsgBox "Código de Retención no existe", vbInformation, "Aviso"
                 End If
              End If
             TxtNumSerieUno = .Fields("Establecimiento")
             TxtNumSerieDos = .Fields("PuntoEmision")
             TxtNumComprobante = .Fields("NumeroComprobantes")
             TxtNumeroC = .Fields("Numero")
             MBFechaEmiV = .Fields("FechaEmision")
             MBFechaRegistroV = .Fields("FechaRegistro")
             TxtBaseImpV = .Fields("BaseImponible")
             TxtBaseImpGravV = .Fields("BaseImpGrav")
             TxtMontoIvaV = .Fields("MontoIva")
             TxtBaseImpoIceV = .Fields("BaseImpIce")
             TxtMontoIceV = .Fields("MontoIce")
             CTP.Text = .Fields("TP")
             TxtNumSerieUno = .Fields("Establecimiento")
             TxtNumSerieDos = .Fields("PuntoEmision")
             TxtIvaBienMonIvaV = .Fields("MontoIvaBienes")
             TxtIvaBienValRetV = .Fields("ValorRetBienes")
             TxtIvaSerMonIvaV = .Fields("MontoIvaServicios")
             TxtIvaSerValRetV = .Fields("ValorRetServicios")
             
             'Busco en el Trans Air para ver si tiene retenciones
             Mifecha = BuscarFecha(MBFechaRegistroV)
             Ln_No = Maximo_De("Asiento_Air", "A_No")
             sSQL = "SELECT TA.*, TC.Concepto " _
                 & "FROM Trans_Air As TA, Tipo_Concepto_Retencion As TC " _
                 & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND TA.Item = '" & NumEmpresa & "' " _
                 & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TA.IdProv = '" & CodProv & "' " _
                 & "AND TC.Fecha_Inicio <= #" & Mifecha & "# " _
                 & "AND TC.Fecha_Final >= #" & Mifecha & "# " _
                 & "AND Tipo_Trans = 'V' " _
                 & "AND TA.Linea_SRI = " & I & " " _
                 & "AND TA.CodRet = TC.Codigo " _
                 & "ORDER BY CodRet,ID "
            'MsgBox sSQL
            SelectAdodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then
               Do While Not AdoAux.Recordset.EOF
                  SetAdoAddNew "Asiento_Air"
                  SetAdoFields "CodRet", AdoAux.Recordset.Fields("CodRet")
                  SetAdoFields "Detalle", AdoAux.Recordset.Fields("Concepto")
                  SetAdoFields "BaseImp", AdoAux.Recordset.Fields("BaseImp")
                  SetAdoFields "Porcentaje", AdoAux.Recordset.Fields("Porcentaje")
                  SetAdoFields "ValRet", AdoAux.Recordset.Fields("ValRet")
                  SetAdoFields "EstabRetencion", AdoAux.Recordset.Fields("EstabRetencion")
                  SetAdoFields "PtoEmiRetencion", AdoAux.Recordset.Fields("PtoEmiRetencion")
                  SetAdoFields "SecRetencion", AdoAux.Recordset.Fields("SecRetencion")
                  SetAdoFields "AutRetencion", AdoAux.Recordset.Fields("AutRetencion")
                  SetAdoFields "FechaEmiRet", AdoAux.Recordset.Fields("Fecha")
                  SetAdoFields "Cta_Retencion", AdoAux.Recordset.Fields("Cta_Retencion")
                  SetAdoFields "EstabFactura", AdoAux.Recordset.Fields("EstabFactura")
                  SetAdoFields "PuntoEmiFactura", AdoAux.Recordset.Fields("PuntoEmiFactura")
                  SetAdoFields "Factura_No", AdoAux.Recordset.Fields("Factura_No")
                  SetAdoFields "IdProv", CodProv
                  SetAdoFields "A_No", Ln_No
                  SetAdoFields "T_No", Trans_No
                  SetAdoFields "Tipo_Trans", "V"
                  SetAdoUpdate
                  Ln_No = Ln_No + 1
                  AdoAux.Recordset.MoveNext
               Loop
            End If
            sSQL = "SELECT * " _
                 & "FROM Asiento_Air " _
                 & "WHERE CodRet <> '.' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND CodigoU = '" & CodigoUsuario & "' " _
                 & "AND T_No = " & Trans_No & " " _
                 & "AND Tipo_Trans = 'V' " _
                 & "ORDER BY CodRet "
            SelectDataGrid DGConceptoAirV, AdoAsientoAir, sSQL

             
             sSQL = "SELECT TA.*, TV.IdProv " _
                  & "FROM Trans_Air As TA, Trans_Ventas As TV " _
                  & "WHERE TA.IdProv = TV.IdProv " _
                  & "AND TA.Linea_SRI = " & I & " " _
                  & "AND TA.Numero = TV.Numero "
             'SelectDataGrid DGConceptoAirV, AdoAsientoAir, sSQL, "Sustento"
            
            'Pongo la Base Imponible
             TxtSumatoria = Val(CCur(TxtBaseImpV)) + Val(CCur(TxtBaseImpGravV))
       End If
   End With
End Sub

Private Sub CTP_GotFocus()
  MarcarTexto CTP
End Sub

Private Sub CTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CTP_LostFocus()
  If IsNumeric(CTP.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     CTP.Text = ""
     CTP.Text = "CD"
     CTP.SetFocus
  End If
End Sub

Private Sub DCClienteV_LostFocus()
  SSTVentas.Tab = 0
  If IsNumeric(DCClienteV.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCClienteV.Text = ""
     Leer_Clientes
     DCClienteV.SetFocus
  Else
     NombreCliente = UCase$(DCClienteV)
     With AdoClientes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & NombreCliente & "' ")
          If Not .EOF Then
             'Busca y captura el codigo de Porcentaje IVA
             CodigoCliente = .Fields("Codigo")
             DireccionCli = .Fields("Direccion")
             CICliente = .Fields("CI_RUC")
             TipoBenef = .Fields("TD")
             LblNumIdentV = CICliente
             LblTD.Caption = TipoBenef
             Carga_TipoComprobantes (TipoBenef)
             DCTipoComprobanteV = "Documentos Autorizados en Ventas excepto ND y NC"
             
             TxtNumComprobante = "000001"
             TxtNumSerietres = "0000001"
             'Aqui despliego el ultimo numero de la Transaccion
             sSQL = "SELECT TOP 1 * " _
                  & "FROM Trans_Ventas " _
                  & "WHERE IdProv = '" & CodigoCliente & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "ORDER BY Secuencial,NumeroComprobantes DESC "
             SelectAdodc AdoAux, sSQL
             With AdoAux.Recordset
              If .RecordCount > 0 Then
                  If .Fields("Secuencial") = 0 Then
                      TxtNumComprobante = .Fields("NumeroComprobantes")
                  Else
                     TxtNumSerietres = .Fields("Secuencial")
                  End If
               End If
             End With
                              
             'Verifica si es Consumidor Final para activar en las fechas los ultimos
             'dias del mes
             If TipoBenef = "C" Then
                MBFechaEmiV = UltimoDiaMes(date)
                MBFechaRegistroV = UltimoDiaMes(date)
                MBFechaEmiV.Enabled = False
                MBFechaRegistroV.Enabled = False
             Else
                MBFechaEmiV.Enabled = True
                MBFechaRegistroV.Enabled = True
             End If
          Else
             FClientesFlash.Show
             Leer_Clientes
          End If
      Else
         FClientesFlash.Show
         Leer_Clientes
      End If
     End With
  End If
End Sub

Private Sub DCConceptoRetV_LostFocus()
  If IsNumeric(DCConceptoRetV.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCConceptoRetV.SetFocus
  Else
     With AdoConceptoRet.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          OP = False
         .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRetV) & "' ")
          Cadena = .Fields("Ingresar_Porcentaje")
          TxtPorRetConAV.Enabled = False
          If Not .EOF Then
             'Verifico el codigo para activar el text e ingrese el porcentaje manualmente
             Select Case Cadena
               Case "N": TxtPorRetConAV = .Fields("Porcentaje")
                    OP = True
               Case "S": TxtPorRetConAV.Enabled = True
             End Select
           Else
              MsgBox "No encontro este código vuelva a buscar"
           End If
      End If
     End With
     TxtBimpConAV = TxtSumatoria
   End If
End Sub

Private Sub DCPorcenIceV_LostFocus()
  If Not IsNumeric(DCPorcenIceV) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIceV = ""
    'Carga_PorcentajeIce
     DCPorcenIceV.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIce = "0"
     With AdoPorIce.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CSng(DCPorcenIceV) & " ")
          If Not .EOF Then CodPorIce = .Fields("Codigo")
       End If
      End With
      Total_IVA = 0
      Total_IVA = CTNumero(TxtBaseImpoIceV, 2)
      TxtMontoIceV = 0
     'Calcula el Porcentaje de Ice
      CalIbMi = Total_IVA * CTNumero(DCPorcenIceV, 2) / 100
      TxtMontoIceV = CalIbMi
  End If

 'Coloca el valor de Monto IVA dependiendo si se activo Bienes o Servicios
  If ChRetB + ChRetS = 0 Then
     TxtIvaBienMonIvaV = TxtMontoIvaV
  End If
  If ChRetB.value <> 0 Then
     TxtIvaBienMonIvaV = TxtMontoIvaV
     TxtIvaSerMonIvaV = 0
  Else
     If ChRetS.value <> 0 Then
        TxtIvaSerMonIvaV = TxtMontoIvaV
        TxtIvaBienMonIvaV = 0
     End If
  End If
End Sub

Private Sub DCPorcenIvaV_LostFocus()
  If Not IsNumeric(DCPorcenIvaV) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIvaV = ""
    'Carga_PorcentajeIva (MBFechaRegis)
     DCPorcenIvaV.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIva = "0"
     With AdoPorIva.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CByte(DCPorcenIvaV) & " ")
          If Not .EOF Then CodPorIva = .Fields("Codigo")
      End If
     End With
     Total_IVA = 0
     Total_IVA = CTNumero(TxtBaseImpGravV, 2)
     'Calcula el Porcentaje de Iva
     CalmIva = (Total_IVA * DCPorcenIvaV) / 100
     TxtMontoIvaV = CalmIva
  End If
End Sub

Private Sub DCPorcenRetenIvaBienV_LostFocus()
  CodRetBien = 0
  If Not IsNumeric(DCPorcenRetenIvaBienV) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenRetenIvaBienV = ""
     Carga_RetencionIvaBienes_Servicios
     DCPorcenRetenIvaBienV.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje retencion Iva Bienes
     With AdoRetIvaBienes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBienV) & " ")
          If Not .EOF Then CodRetBien = .Fields("Codigo")
      Else
         MsgBox "Código incorrecto", vbInformation, "Aviso"
      End If
     End With
     Total_IVA = CTNumero(TxtIvaBienMonIvaV, 2)
    'Calcula la retencion Iva Bienes
     CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBienV)) / 100
     TxtIvaBienValRetV = CalIbMi
  End If
  TxtIvaSerMonIvaV = Format(CTNumero(TxtMontoIvaV, 2) - CTNumero(TxtIvaBienMonIvaV, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServV_LostFocus()
  CodRetServ = 0
  'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
  If DCPorcenRetenIvaServV = "70/100" Then
     Ct = "Si"
     TxtIvaSerValRetV.Text = ""
     TxtIvaSerValRetV.Enabled = True
     TxtIvaSerValRetV.SetFocus
  Else
     If Not IsNumeric(DCPorcenRetenIvaServV) Then
        MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
        DCPorcenRetenIvaServV = ""
        Carga_RetencionIvaBienes_Servicios
        DCPorcenRetenIvaServV.SetFocus
     End If
  End If
    
 'Busca captura el codigo de Porcentaje retencion Iva Servicios
  With AdoRetIvaServicios.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServV) & " ")
       If Not .EOF Then CodRetServ = .Fields("Codigo")
   Else
      MsgBox "Código Incorrecto", vbInformation, "Aviso"
   End If
  End With
  Ct = "No"
  Total_IVA = 0
  Total_IVA = CTNumero(TxtIvaSerMonIvaV, 2)
  If DCPorcenRetenIvaServV = "70/100" Then
  Else
     CalIsMi = (Total_IVA * CInt(DCPorcenRetenIvaServV)) / 100
     TxtIvaSerValRetV = CalIsMi
     TxtIvaSerValRetV.Enabled = False
  End If
  SSTVentas.Tab = 0
  SSTVentas.SetFocus
End Sub

Private Sub DCTipoComprobanteV_LostFocus()
  If IsNumeric(DCTipoComprobanteV) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCTipoComprobanteV.Text = ""
     DCTipoComprobanteV.SetFocus
     Captura_TipoComprobanteV
  Else
     If DCTipoComprobanteV <> "" Then
        Captura_TipoComprobanteV
     End If
  End If
End Sub

Private Sub DGConceptoAirV_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     Titulo = "Aviso"
     Mensajes = "Desea Eliminar la Retención"
     If BoxMensaje = vbYes Then
        With AdoAsientoAir.Recordset
         If .RecordCount > 0 Then
             Codigo = .Fields("CodRet")
             No_Desde = .Fields("SecRetencion")
             Mifecha = .Fields("FechaEmiRet")
             Codigo1 = .Fields("AutRetencion")
             J = .Fields("A_No")
             sSQL = "DELETE * " _
                  & "FROM Asiento_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND IdProv = '" & CodigoCliente & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND Tipo_Trans = 'V' " _
                  & "AND A_No = " & J & " " _
                  & "AND CodRet = '" & Codigo & "' "
             ConectarAdoExecute sSQL
         End If
         AdoAsientoAir.Refresh
        End With
        Calculo_Sumatoria
     End If
  End If
End Sub

Private Sub MBFechaEmiV_GotFocus()
  MarcarTexto MBFechaEmiV
End Sub

Private Sub MBFechaEmiV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiV_LostFocus()
  FechaValida MBFechaEmiV
  'Controla que la Fecha de Emisiòn este entre 31/01/2000 en adelante
  If CFechaLong(MBFechaEmiV) < CFechaLong("31/01/2000") Then
     MsgBox "La Fecha de Emisión debe ser mayor que 31/01/2000", vbInformation, "Aviso"
     MBFechaEmiV = "31/01/2000"
     MBFechaEmiV.SetFocus
  End If
  MBFechaRegistroV = MBFechaEmiV
  Carga_ConceptosRetencion MBFechaRegistroV
End Sub

Private Sub MBFechaRegistroV_GotFocus()
  MarcarTexto MBFechaRegistroV
End Sub

Private Sub MBFechaRegistroV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegistroV_LostFocus()
  FechaValida MBFechaRegistroV
  'Controla que la Fecha de Emisiòn este entre 31/01/2000 en adelante
  If CFechaLong(MBFechaRegistroV) < CFechaLong("31/01/2000") Then
     MsgBox "La Fecha de Registro debe ser mayor que 31/01/2000", vbInformation, "Aviso"
     MBFechaRegistroV = "31/01/2000"
     MBFechaRegistroV.SetFocus
  End If
  'Carga el porcentaje de IVA
  Carga_ConceptosRetencion MBFechaRegistroV
  Carga_RetencionIvaBienes_Servicios
End Sub

Private Sub OpcIvaNo_LostFocus()
  If OpcIvaNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcIvaSi_LostFocus()
  If OpcIvaSi.value = True Then ValorP = "S"
End Sub

Private Sub OpcRetNo_Click()
  If OpcRetNo.value = True Then ValorR = "S"
End Sub

Private Sub OpcRetSi_LostFocus()
  If OpcRetSi.value = True Then ValorR = "S"
End Sub

Private Sub SSTVentas_Click(PreviousTab As Integer)
  If PreviousTab = 1 Then DCClienteV.SetFocus
End Sub

Private Sub TxtBaseImpGravV_GotFocus()
  MarcarTexto TxtBaseImpGravV
End Sub

Private Sub TxtBaseImpGravV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpGravV_LostFocus()
  TextoValido TxtBaseImpGravV, True, , 0
End Sub

Private Sub TxtBaseImpoIceV_GotFocus()
  MarcarTexto TxtBaseImpoIceV
End Sub

Private Sub TxtBaseImpoIceV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIceV_LostFocus()
  TextoValido TxtBaseImpoIceV, True, , 0
End Sub

Private Sub TxtBaseImpV_GotFocus()
  MarcarTexto TxtBaseImpV
End Sub

Private Sub TxtBaseImpV_LostFocus()
  TextoValido TxtBaseImpV, True, , 0
  FechaValida MBFechaRegistroV
End Sub

Private Sub TxtBaseImpV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  Ln_No = 1
  Ln_SRI = -1
  Carga_RetencionIvaBienes_Servicios
  CMes.Clear
  sSQL = Listar_Meses
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CMes.AddItem .Fields("Dia_Mes")
          CMes.Tag = .Fields("No_D_M")
         .MoveNext
       Loop
   End If
  End With
  CMes.Text = CMes.List(0)
  For I = Year(FechaSistema) To 2000 Step -1
      CAño.AddItem Format(I, "0000")
  Next I
  CAño.Text = CAño.List(0)
  Carga_Datos_Iniciales MBFecha, Nuevo
End Sub

Private Sub Form_Load()
  CentrarForm FVentas
  ConectarAdodc AdoSustento
  ConectarAdodc AdoTipoIdentificacion
  ConectarAdodc AdoTipoComprobante
  ConectarAdodc AdoRetIvaBienes
  ConectarAdodc AdoRetIvaServicios
  ConectarAdodc AdoPorIce
  ConectarAdodc AdoPorIva
  ConectarAdodc AdoConceptoRet
  ConectarAdodc AdoAsientoAir
  ConectarAdodc AdoAsientoVentas
  ConectarAdodc AdoTransAir
  ConectarAdodc AdoTransVentas
  ConectarAdodc AdoClientes
  ConectarAdodc AdoRetFuente
  ConectarAdodc AdoRetIvaSerCC
  ConectarAdodc AdoRetIvaBienesCC
  ConectarAdodc AdoAux
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <>  '.' " _
       & "AND TD <>  'E' " _
       & "ORDER BY  Cliente "
  SelectDBCombo DCClienteV, AdoClientes, sSQL, "Cliente"
End Sub

Private Sub TxtBimpConAV_GotFocus()
  MarcarTexto TxtBimpConAV
End Sub

Private Sub TxtBimpConAV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBimpConAV_LostFocus()
  RatonNormal
  TextoValido TxtBimpConAV, True, , 0
  TextoValido TxtSumatoria, True, , 0
  'Valida que la base imponible no sea mayor que la BIG y la BIcero
  If CTNumero(TxtBimpConAV, 2) > CTNumero(TxtSumatoria, 2) Then
     MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
     & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
     TxtBimpConAV.Text = 0
     TxtBimpConAV.SetFocus
  Else
     If OP Then
        TxtValConAV = (CTNumero(TxtBimpConAV, 2) * CTNumero(TxtPorRetConAV, 2)) / 100
        Insertar_AsientoAir
     Else
        TxtPorRetConAV.SetFocus
     End If
  End If
End Sub

Sub Insertar_AsientoAir()
  'Selecciona el numero mayor para continuar la secuencia en el
  'campo T_No y A_No
  sSQL = "SELECT TOP 1 * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY A_No DESC "
  SelectAdodc AdoAsientoAir, sSQL
  If AdoAsientoAir.Recordset.RecordCount > 0 Then Ln_No = AdoAsientoAir.Recordset.Fields("A_No") + 1
     If CTNumero(TxtBimpConAV, 2) > 0 Then
        RatonReloj
        If Val(TxtNumComprobante) > 1 Then
           Factura_No = Val(TxtNumComprobante)
        Else
           Factura_No = Val(TxtNumSerietres)
        End If
        Espizq = SinEspaciosIzq(DCConceptoRetV)
        Espder = Trim$(Mid$(DCConceptoRetV, Len(Espizq) + 3, Len(DCConceptoRetV)))
        SetAdoAddNew "Asiento_Air"
        SetAdoFields "CodRet", Espizq
        SetAdoFields "Detalle", Espder
        SetAdoFields "BaseImp", CTNumero(TxtBimpConAV, 2)
        SetAdoFields "Porcentaje", CTNumero(TxtPorRetConAV, 2) / 100
        SetAdoFields "ValRet", CTNumero(TxtValConAV, 2)
        SetAdoFields "EstabRetencion", TxtNumUnoComRetV
        SetAdoFields "PtoEmiRetencion", TxtNumDosComRetV
        SetAdoFields "SecRetencion", CTNumero(TxtNumTresComRetV)
        SetAdoFields "AutRetencion", TxtNumUnoAutComRetV
        SetAdoFields "FechaEmiRet", MBFechaRegistroV
        SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
        SetAdoFields "EstabFactura", TxtNumSerieUno
        SetAdoFields "PuntoEmiFactura", TxtNumSerieDos
        SetAdoFields "Factura_No", CTNumero(TxtNumSerietres)
        SetAdoFields "IdProv", CodigoCliente
        SetAdoFields "A_No", Ln_No
        SetAdoFields "T_No", Trans_No
        SetAdoFields "Tipo_Trans", "V"
        SetAdoUpdate
        Ln_No = Ln_No + 1
           
        'Despliega los datos en el DataGrid
        sSQL = "SELECT * " _
             & "FROM Asiento_Air " _
             & "WHERE CodRet <> '.' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "AND T_No = " & Trans_No & " " _
             & "AND Tipo_Trans = 'V' " _
             & "ORDER BY CodRet "
        SelectDataGrid DGConceptoAirV, AdoAsientoAir, sSQL
      
        'Se situa en el combo de retención AIR
        If ChRetF.Visible Then DCRetFuente.SetFocus Else TxtNumUnoComRetV.SetFocus
          
        'Realiza la Sumatoria de las Retenciones
        ac = ac + TxtValConAV
        TxtTotalReten = ac
   End If
   RatonNormal
End Sub

Private Sub TxtIvaBienMonIvaV_GotFocus()
  MarcarTexto TxtIvaBienMonIvaV
End Sub

Private Sub TxtIvaBienMonIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIvaBienMonIvaV_LostFocus()
  TextoValido TxtIvaBienMonIvaV, True, , 2
End Sub

Private Sub TxtIvaBienValRetV_GotFocus()
  MarcarTexto TxtIvaBienValRetV
End Sub

Private Sub TxtIvaBienValRetV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIvaBienValRetV_LostFocus()
  TextoValido TxtIvaBienValRetV, True, , 0
End Sub

Private Sub TxtIvaSerMonIvaV_GotFocus()
  TextoValido TxtIvaSerMonIvaV, True, , 2
  MarcarTexto TxtIvaSerMonIvaV
End Sub

Private Sub TxtIvaSerMonIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIvaSerMonIvaV_LostFocus()
  TextoValido TxtIvaSerMonIvaV, True, , 0
    
  'Verifica el Monto Iva Servicios
  If CDbl(TxtIvaBienMonIvaV) + CDbl(TxtIvaSerMonIvaV) > CDbl(TxtMontoIvaV) Then
     MsgBox "Monto IVA Servicios no puede ser > que Monto IVA", vbInformation, "Aviso de Ventas"
     TxtIvaSerMonIvaV.Text = ""
     TxtIvaSerMonIvaV.SetFocus
  End If
End Sub

Private Sub TxtMontoIvaV_GotFocus()
  MarcarTexto TxtMontoIvaV
End Sub

Private Sub TxtMontoIvaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIvaV_LostFocus()
  TextoValido TxtMontoIvaV, True, , 2
End Sub

Private Sub TxtNumComprobante_GotFocus()
  MarcarTexto TxtNumComprobante
End Sub

Private Sub TxtNumComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumComprobante_LostFocus()
  TextoValido TxtNumComprobante, True, , 0
  If Val(TxtNumComprobante) <= 0 Then TxtNumComprobante = "0000001"
     TxtNumComprobante = Format(Val(CCur(TxtNumComprobante)), "0000000")
    'Verifico si es uno o más comprobantes
    If CLng(TxtNumComprobante) <> 1 And TipoBenef <> "C" Then
       MBFechaEmiV.SetFocus
       TxtNumSerietres = "0000001"
       TxtNumSerietres.Enabled = False
    Else
       TxtNumSerietres.Enabled = True
       TxtNumSerieUno.SetFocus
    End If
End Sub

Private Sub TxtNumDosComRetV_GotFocus()
  MarcarTexto TxtNumDosComRetV
End Sub

Private Sub TxtNumDosComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumDosComRetV_LostFocus()
  TextoValido TxtNumDosComRetV, True, , 0
  If Val(TxtNumDosComRetV) <= 0 Then TxtNumDosComRetV = "001"
  TxtNumDosComRetV = Format(Val(TxtNumDosComRetV), "000")
End Sub

Private Sub TxtNumeroC_GotFocus()
  MarcarTexto TxtNumeroC
End Sub

Private Sub TxtNumeroC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumeroC_LostFocus()
  If Not IsNumeric(TxtNumeroC.Text) Then
     MsgBox "No ingrese caracteres alfabéticos. Vuelva a ingresar.", vbInformation, "Aviso"
     TxtNumeroC.Text = ""
     TxtNumeroC.SetFocus
  End If
  If TxtNumeroC.Text = "." Or TxtNumeroC.Text = " " Then
     MsgBox "Ingrese el Número de Comprobante", vbInformation, "Aviso"
  End If
  If Val(TxtNumeroC) <= 0 Then TxtNumeroC = "0"
     TxtNumeroC = Format(CTNumero(TxtNumeroC), String(7, "0"))
End Sub

Private Sub TxtNumSerieDos_GotFocus()
    MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
  TextoValido TxtNumSerieDos, True, , 0
  If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
  TxtNumSerieDos = Format(Val(TxtNumSerieDos), "000")
End Sub

Private Sub TxtNumSerietres_GotFocus()
  MarcarTexto TxtNumSerietres
End Sub

Private Sub TxtNumSerietres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres_LostFocus()
  If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = "000000001"
  TxtNumSerietres = Format(Val(Round(TxtNumSerietres)), "000000000")
End Sub

Private Sub TxtNumSerieUno_GotFocus()
  MarcarTexto TxtNumSerieUno
End Sub

Private Sub TxtNumSerieUno_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUno_LostFocus()
  TextoValido TxtNumSerieUno, True, , 0
  If Val(TxtNumSerieUno) <= 0 Then TxtNumSerieUno = "001"
  TxtNumSerieUno = Format(Val(TxtNumSerieUno), "000")
End Sub

Private Sub TxtNumTresComRetV_GotFocus()
  MarcarTexto TxtNumTresComRetV
End Sub

Private Sub TxtNumTresComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRetV_LostFocus()
  If Val(TxtNumTresComRetV) <= 0 Then TxtNumTresComRetV = "00000001"
  TxtNumTresComRetV = Format(Val(Round(TxtNumTresComRetV)), "000000000")
    
  'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
  'TxtSumatoria = CDbl(TxtIvaBienMonIvaV) + CDbl(TxtIvaSerMonIvaV) + CDbl(TxtBaseImpV)
  TxtSumatoria = Val(CCur(TxtBaseImpV)) + Val(CCur(TxtBaseImpGravV))
End Sub

Private Sub TxtNumUnoAutComRetV_GotFocus()
  MarcarTexto TxtNumUnoAutComRetV
End Sub

Private Sub TxtNumUnoAutComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRetV_LostFocus()
  If Val(TxtNumUnoAutComRetV) <= 0 Then TxtNumUnoAutComRetV = "0"
  TxtNumUnoAutComRetV = Format(Val(Round(TxtNumUnoAutComRetV)), String(10, "0"))
End Sub

Private Sub TxtNumUnoComRetV_GotFocus()
  MarcarTexto TxtNumUnoComRetV
End Sub

Private Sub TxtNumUnoComRetV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRetV_LostFocus()
  TextoValido TxtNumUnoComRetV, True, , 0
  If Val(TxtNumUnoComRetV) <= 0 Then TxtNumUnoComRetV = "001"
  TxtNumUnoComRetV = Format(Val(TxtNumUnoComRetV), "000")
End Sub

Public Sub Captura_TipoComprobanteV()
 'Captura lo que tiene el Combo de Tipo de Comprobante
  Captc = SinEspaciosIzq(DCTipoComprobanteV.Text)
  Cap1 = Trim$(DCTipoComprobanteV.Text)
  cod = 0
 'Busca que sea igual a la Descripcion
  With AdoTipoComprobante.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & Cap1 & "' ")
       If Not .EOF Then
          cod = .Fields("Tipo_Comprobante_Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
       End If
    End If
  End With
End Sub

Public Sub Carga_RetencionIvaBienes_Servicios()
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Bienes <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDBCombo DCPorcenRetenIvaBienV, AdoRetIvaBienes, sSQL, "Porc"
   
   sSQL = "SELECT * " _
        & "FROM Tabla_Por_IVA " _
        & "WHERE Servicios <> " & Val(adFalse) & " " _
        & "ORDER BY Porc "
   SelectDBCombo DCPorcenRetenIvaServV, AdoRetIvaServicios, sSQL, "Porc"
End Sub

Public Sub Carga_TipoComprobantes(CargaTC As String)
  If CargaTC = "O" Then CargaTC = "P"
  'Carga en el combo los tipos de comprobantes de acuerdo a la Identificacion
   sSQL = "SELECT CTT.Identificacion,CTT.Tipo_Trans,TC.* " _
        & "FROM Tabla_Tributaria As CTT, Tipo_Comprobante As TC " _
        & "WHERE CTT.Identificacion = '" & CargaTC & "' " _
        & "AND CTT.Tipo_Trans = 'V' " _
        & "AND CTT.Tipo_Comprobante_Codigo = TC.Tipo_Comprobante_Codigo " _
        & "ORDER BY TC.Tipo_Comprobante_Codigo "
   SelectDBCombo DCTipoComprobanteV, AdoTipoComprobante, sSQL, "Descripcion"
   DCTipoComprobanteV = "Documentos Autorizados en Ventas excepto ND y NC"
End Sub

Public Sub Limpiar_Controles()
  ac = 0
  DCClienteV.Text = ""
  TxtNumeroC.Text = ""
  DCRetIBienes.Visible = False
  DCRetISer.Visible = False
  ChRetB.value = False
  ChRetS.value = False
  ChRetF.value = False
  LblNumIdentV.Caption = ""
  LblTD.Caption = ""
  OpcIvaNo.value = True
  OpcRetNo.value = True
  DCTipoComprobanteV.Text = ""
  TxtNumComprobante.Text = ""
  FechaValida MBFechaEmiV
  FechaValida MBFechaRegistroV
  TxtBaseImpV.Text = ""
  TxtBaseImpGravV.Text = ""
  TxtBaseImpoIceV.Text = ""
  DCPorcenIvaV.Text = ""
  TxtMontoIvaV.Text = ""
  DCPorcenIceV.Text = ""
  TxtMontoIceV.Text = ""
  TxtIvaBienMonIvaV.Text = ""
  DCPorcenRetenIvaBienV.Text = ""
  TxtIvaBienValRetV.Text = ""
  TxtIvaSerMonIvaV.Text = ""
  DCPorcenRetenIvaServV.Text = ""
  TxtIvaSerValRetV.Text = ""
  TxtNumUnoComRetV.Text = ""
  TxtNumDosComRetV.Text = ""
  TxtNumTresComRetV.Text = ""
  TxtNumUnoAutComRetV.Text = ""
  TxtSumatoria.Text = ""
  DCConceptoRetV.Text = ""
  TxtBimpConAV.Text = ""
  TxtPorRetConAV.Text = ""
  TxtValConAV.Text = ""
  TxtTotalReten.Text = ""
  CTP.Clear
  CTP.AddItem "CE"
  CTP.AddItem "CI"
  CTP.AddItem "CD"
  CTP.Text = "CE"
   
  'Limpia la grilla
  'Borra Asiento Air
  sSQL = "DELETE * " _
       & "FROM Asiento_Air " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND Tipo_Trans = 'V' "
 ConectarAdoExecute sSQL
End Sub

Public Sub Calculo_Sumatoria()
Dim SumaReten As Currency
  SumaReten = 0
  With AdoAsientoAir.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaReten = SumaReten + .Fields("ValRet")
         .MoveNext
       Loop
   End If
  End With
  TxtTotalReten = Format(SumaReten, "#,##0.00")
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    ' Label16.Caption = "No.Autoriz.:" & Autorizacion
    Encerar_Var
    Limpiar_Controles
   'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
    Listar_Air
   'Carga en el Data Combo los Clientes con su RUC
    Leer_Clientes
   'Carga la Tabla de Conceptos Retencion al DataCombo
    Carga_ConceptosRetencion MBFechaRegistroV
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CF' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
    With AdoRetFuente.Recordset
     If .RecordCount > 0 Then Rf = 1 Else Rf = 0
    End With
  
   'Carga los Conceptos de retención IVA Servicios al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetISer, AdoRetIvaSerCC, sSQL, "Cuentas"
    With AdoRetIvaSerCC.Recordset
     If .RecordCount > 0 Then Rs = 1 Else Rs = 0
    End With
    
    'Carga los Conceptos de retención IVA Bienes al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'CI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetIBienes, AdoRetIvaBienesCC, sSQL, "Cuentas"
    With AdoRetIvaBienesCC.Recordset
     If .RecordCount > 0 Then Rb = 1 Else Rb = 0
    End With
    
   'Si es Nuevo ingresa por aqui
    ChRetF.Visible = True
    DCRetFuente.Visible = True
    FrmRetencion.Visible = True
    'LblMensaje.Visible = False
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And Rs And Rb) = 0 Then
           ChRetF.Visible = False
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           'LblMensaje.Visible = True
           Activar_BS
           CTP.SetFocus
       Else
           ChRetB.SetFocus
       End If
    Else
      'Si es Modificación viene por aca
       CMes.Visible = True
       CAño.Visible = True
       CMes.Text = MesesLetras(Month(FechaSistema))
       CAño.Text = Year(FechaSistema)
       Modificacion_AT "V", CModificacion, CMes, CAño
       CModificacion.Visible = True
       CAño.SetFocus
     End If
End Sub

Public Sub Grabacion()
   'Selecciona el numero mayor para continuar la secuencia en el
   'campo T_No y A_No
   'Grabo en el Asiento_Ventas e implicito Asiento_Air
    Captura_TipoComprobanteV
    SetAdoAddNew "Asiento_Ventas"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "FechaRegistro", MBFechaRegistroV
    SetAdoFields "FechaEmision", MBFechaEmiV
    SetAdoFields "BaseImponible", CTNumero(TxtBaseImpV, 2)
    SetAdoFields "IvaPresuntivo", ValorP
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    If Val(TxtNumComprobante) > 1 Then
       SetAdoFields "NumeroComprobantes", CTNumero(TxtNumComprobante)
       SetAdoFields "Secuencial", 0
    Else
       SetAdoFields "Secuencial", CTNumero(TxtNumSerietres)
       SetAdoFields "NumeroComprobantes", 1
    End If
    
    SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpGravV, 2)
    SetAdoFields "PorcentajeIva", CodPorIva
    SetAdoFields "MontoIva", CTNumero(TxtMontoIvaV, 2)
    SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIceV, 2)
    SetAdoFields "PorcentajeIce", CodPorIce
    SetAdoFields "MontoIce", CTNumero(TxtMontoIceV, 2)
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIvaV, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRetV, 2)
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIvaV, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRetV, 2)
    SetAdoFields "RetPresuntiva", ValorR
   'Verifico si activaron los checks
    If ChRetB = 1 Then
       SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    Else
       SetAdoFields "Cta_Bienes", "."
    End If
    If ChRetS = 1 Then
       SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
    Else
       SetAdoFields "Cta_Servicio", "."
    End If
    
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBienV
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIvaV, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRetV, 2)
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServV
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIvaV, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRetV, 2)
    SetAdoFields "A_No", 0
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
         
   'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = Maximo_De("Trans_Ventas", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT * " _
         & "FROM Asiento_Ventas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No DESC "
    SelectAdodc AdoAsientoVentas, sSQL
    With AdoAsientoVentas.Recordset
     If .RecordCount > 0 Then
         FechaTexto = .Fields("FechaRegistro")
         SetAdoAddNew "Trans_Ventas"
         SetAdoFields "T", Normal
         SetAdoFields "IdProv", CodigoCliente
         SetAdoFields "TipoComprobante", .Fields("TipoComprobante")
         SetAdoFields "FechaRegistro", .Fields("FechaRegistro")
         SetAdoFields "FechaEmision", .Fields("FechaEmision")
         SetAdoFields "Establecimiento", .Fields("Establecimiento")
         SetAdoFields "PuntoEmision", .Fields("PuntoEmision")
         SetAdoFields "Secuencial", .Fields("Secuencial")
         SetAdoFields "NumeroComprobantes", .Fields("NumeroComprobantes")
         SetAdoFields "BaseImponible", .Fields("BaseImponible")
         SetAdoFields "IvaPresuntivo", .Fields("IvaPresuntivo")
         SetAdoFields "BaseImpGrav", .Fields("BaseImpGrav")
         SetAdoFields "PorcentajeIva", .Fields("PorcentajeIva")
         SetAdoFields "MontoIva", .Fields("MontoIva")
         SetAdoFields "BaseImpIce", .Fields("BaseImpIce")
         SetAdoFields "PorcentajeIce", .Fields("PorcentajeIce")
         SetAdoFields "MontoIce", .Fields("MontoIce")
         SetAdoFields "MontoIvaBienes", .Fields("MontoIvaBienes")
         SetAdoFields "PorRetBienes", .Fields("PorRetBienes")
         SetAdoFields "ValorRetBienes", .Fields("ValorRetBienes")
         SetAdoFields "MontoIvaServicios", .Fields("MontoIvaServicios")
         SetAdoFields "PorRetServicios", .Fields("PorRetServicios")
         SetAdoFields "ValorRetServicios", .Fields("ValorRetServicios")
         SetAdoFields "RetPresuntiva", .Fields("RetPresuntiva")
         SetAdoFields "TP", CTP
         SetAdoFields "Numero", TxtNumeroC
         SetAdoFields "Fecha", MBFechaRegistroV
         SetAdoFields "Porc_Bienes", .Fields("Porc_Bienes")
         SetAdoFields "Porc_Servicios", .Fields("Porc_Servicios")
         SetAdoFields "Cta_Servicio", .Fields("Cta_Servicio")
         SetAdoFields "Cta_Bienes", .Fields("Cta_Bienes")
         SetAdoFields "ID", ID_Trans
         SetAdoFields "Linea_SRI", 0
         SetAdoUpdate
     End If
    End With
    
   'Selecciona el numero mayor para continuar la secuencia en el
    ID_Trans = Maximo_De("Trans_Air", "ID")
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'V' " _
         & "ORDER BY A_No "
    SelectAdodc AdoTransAir, sSQL
    With AdoAsientoAir.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            SetAdoAddNew "Trans_Air"
            SetAdoFields "T", Normal
            SetAdoFields "CodRet", .Fields("CodRet")
            SetAdoFields "BaseImp", .Fields("BaseImp")
            SetAdoFields "Porcentaje", .Fields("Porcentaje")
            SetAdoFields "ValRet", .Fields("ValRet")
            SetAdoFields "EstabRetencion", .Fields("EstabRetencion")
            SetAdoFields "PtoEmiRetencion", .Fields("PtoEmiRetencion")
            SetAdoFields "SecRetencion", .Fields("SecRetencion")
            SetAdoFields "AutRetencion", .Fields("AutRetencion")
            SetAdoFields "EstabFactura", .Fields("EstabFactura")
            SetAdoFields "PuntoEmiFactura", .Fields("PuntoEmiFactura")
            SetAdoFields "Factura_No", .Fields("Factura_No")
            SetAdoFields "Tipo_Trans", .Fields("Tipo_Trans")
            SetAdoFields "Fecha", FechaTexto
            SetAdoFields "IdProv", CodigoCliente
            SetAdoFields "TP", CTP
            SetAdoFields "Numero", TxtNumeroC
            SetAdoFields "Cta_Retencion", .Fields("Cta_Retencion")
            SetAdoFields "ID", ID_Trans
            SetAdoFields "Linea_SRI", 0
            SetAdoUpdate
           .MoveNext
         Loop
       Else
          'MsgBox "No existen datos para ingresar Trans Air", vbCritical, "Aviso ventas"
     End If
    End With

End Sub

Public Sub Habilita_Controles()
    'Habilito los controles para la modificacion
    CModificacion.Enabled = True
    SSTVentas.Enabled = True
    DCClienteV.Enabled = True
    CmdGrabar.Enabled = True
    FrmTipoComprob.Enabled = True
    FrmRetencion.Enabled = True
    CMes.Visible = True
    CAño.Visible = True
End Sub

Public Sub Deshabilita_Controles()
    'Deshabilito los controles para la modificacion
    CModificacion.Enabled = False
    SSTVentas.Enabled = False
    DCClienteV.Enabled = False
    CmdGrabar.Enabled = False
    FrmTipoComprob.Enabled = False
    FrmRetencion.Enabled = False
    CMes.Visible = False
    CAño.Visible = False
End Sub

Public Sub Activar_BS()
    TxtIvaBienMonIvaV.Enabled = True
    DCPorcenRetenIvaBienV.Enabled = True
    TxtIvaBienValRetV.Enabled = True
    TxtIvaSerMonIvaV.Enabled = True
    DCPorcenRetenIvaServV.Enabled = True
    TxtIvaSerValRetV.Enabled = True
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Ventas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'V' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Presentamos la malla Asiento Air
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'V' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAirV, AdoAsientoAir, sSQL
End Sub

Public Sub Carga_ConceptosRetencion(MBFecha As String)
Dim FechaCodAir As String
  FechaCodAir = BuscarFecha(MBFecha)
 'Carga la Tabla de Porcentaje Iva
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE IVA <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenIvaV, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDBCombo DCPorcenIceV, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConceptoRetV, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  DCConceptoRetV = "329 - Por Otros Servicios (N)"
End Sub

Private Sub TxtPorRetConAV_GotFocus()
  MarcarTexto TxtPorRetConAV
End Sub

Private Sub TxtPorRetConAV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter TxtPorRetConAV
End Sub

Private Sub TxtPorRetConAV_LostFocus()
  If OP = False Then
     TxtValConAV = (CTNumero(TxtBimpConAV, 2) * CTNumero(TxtPorRetConAV, 2)) / 100
     Insertar_AsientoAir
  End If
End Sub

Public Sub Encerar_Var()
  ac = 0
  cod = 0
  Ln_No = 0
  DCPorcenIceV = 0
  DCPorcenRetenIvaBienV = 0
  DCPorcenRetenIvaServV = 0
  CodPorIce = "0"
  CodPorIva = "0"
  CodRetBien = "0"
  CodRetServ = "0"
End Sub

