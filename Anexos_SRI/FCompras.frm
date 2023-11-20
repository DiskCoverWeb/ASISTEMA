VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPRAS"
   ClientHeight    =   7440
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10590
   ForeColor       =   &H8000000F&
   Icon            =   "FCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10590
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "| CORRECCION DE FORMULARIOS DE COMPRAS |"
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
      TabIndex        =   109
      Top             =   0
      Width           =   10410
      Begin VB.ComboBox CAño 
         BackColor       =   &H0080FFFF&
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
         ItemData        =   "FCompras.frx":0696
         Left            =   105
         List            =   "FCompras.frx":0698
         TabIndex        =   112
         Text            =   "2000"
         Top             =   210
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox CModificacion 
         BackColor       =   &H0080FFFF&
         DataSource      =   "AdoTransCompras"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   111
         Top             =   210
         Visible         =   0   'False
         Width           =   7755
      End
      Begin VB.ComboBox CMes 
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   110
         Text            =   "Combo1"
         Top             =   210
         Visible         =   0   'False
         Width           =   1416
      End
   End
   Begin VB.Frame FrmRetencion 
      Caption         =   "RETENCIONES DE IVA POR:"
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
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   7365
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FCompras.frx":069A
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   525
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   960
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
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   525
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FCompras.frx":06B7
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
   End
   Begin VB.Frame FrmTipoComprob 
      Height          =   960
      Left            =   7455
      TabIndex        =   6
      Top             =   420
      Width           =   1905
      Begin VB.TextBox TxtNumeroC 
         Alignment       =   1  'Right Justify
         Height          =   336
         Left            =   840
         MaxLength       =   9
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "FCompras.frx":06D7
         ToolTipText     =   "En este campo se debe ingresar el número del comprobante, el cual no excedera los siete caracteres"
         Top             =   525
         Width           =   975
      End
      Begin VB.ComboBox CTP 
         Height          =   315
         Left            =   105
         TabIndex        =   7
         ToolTipText     =   "En este combo se despliega una lista con lo stipos de comprobantes existentes tales como: Comprobante Diario, Ingreso o Egreso"
         Top             =   525
         Width           =   660
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
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
         Left            =   840
         TabIndex        =   72
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TP"
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
         TabIndex        =   71
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   750
      Left            =   9450
      Picture         =   "FCompras.frx":06DB
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Grabar"
      Top             =   324
      Width           =   960
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Salir"
      Height          =   765
      Left            =   9450
      Picture         =   "FCompras.frx":09E5
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Salir"
      Top             =   1155
      Width           =   990
   End
   Begin TabDlg.SSTab SSTCompras 
      Height          =   5400
      Left            =   108
      TabIndex        =   0
      Top             =   1944
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9525
      _Version        =   393216
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- Comprobante de Compra"
      TabPicture(0)   =   "FCompras.frx":0E27
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DCSustento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OpcSi"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OpcNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraDctoModificado"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdAir"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&2.- Conceptos AIR"
      TabPicture(1)   =   "FCompras.frx":0E43
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3.- Partidos Políticos"
      TabPicture(2)   =   "FCompras.frx":0E5F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton CmdAir 
         Caption         =   "&AIR"
         Height          =   444
         Left            =   9660
         Picture         =   "FCompras.frx":0E7B
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   420
         Width           =   552
      End
      Begin VB.Frame Frame8 
         Caption         =   "SOLO PARTIDOS POLITICOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74895
         TabIndex        =   62
         Top             =   420
         Width           =   10050
         Begin VB.TextBox TxtMonTitGrat 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   65
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no oneroso para el informante"
            Top             =   2730
            Width           =   1905
         End
         Begin VB.TextBox TxtMonTitOner 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   64
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no gratuito para el informante"
            Top             =   1995
            Width           =   1905
         End
         Begin VB.TextBox TxtNumConParPol 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            MaxLength       =   10
            TabIndex        =   63
            Text            =   "0000000000"
            ToolTipText     =   $"FCompras.frx":13A1
            Top             =   1260
            Width           =   1905
         End
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO DEL CONTRATO"
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
            TabIndex        =   99
            Top             =   2730
            Width           =   4635
         End
         Begin VB.Label Label38 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO TITULO ONEROSO"
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
            TabIndex        =   98
            Top             =   1995
            Width           =   4635
         End
         Begin VB.Label Label37 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NUMERO DE CONTRATO DEL PARTIDO POLITICO"
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
            TabIndex        =   97
            Top             =   1260
            Width           =   4635
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
         Height          =   1485
         Left            =   105
         TabIndex        =   27
         Top             =   2940
         Width           =   4950
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   29
            Text            =   "0.00"
            ToolTipText     =   "Este valor se calcula automaticamente, es el resultado de aplicarle un porcentaje IVA a la Base Imponible gravada"
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtMontoIce 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3465
            TabIndex        =   31
            Top             =   945
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FCompras.frx":145C
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   28
            ToolTipText     =   $"FCompras.frx":1474
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FCompras.frx":1506
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   30
            ToolTipText     =   $"FCompras.frx":151E
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            TabIndex        =   86
            Top             =   945
            Width           =   1485
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
            TabIndex        =   85
            Top             =   420
            Width           =   1485
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
            TabIndex        =   84
            Top             =   945
            Width           =   855
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
            TabIndex        =   83
            Top             =   420
            Width           =   855
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
         Height          =   1485
         Left            =   5145
         TabIndex        =   32
         Top             =   2940
         Width           =   5055
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServ 
            Bindings        =   "FCompras.frx":15AF
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3150
            TabIndex        =   37
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBien 
            Bindings        =   "FCompras.frx":15D0
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1470
            TabIndex        =   34
            ToolTipText     =   $"FCompras.frx":15EE
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.TextBox TxtIvaBienMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "FCompras.frx":167A
            ToolTipText     =   $"FCompras.frx":167F
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaBienValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            TabIndex        =   35
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            MultiLine       =   -1  'True
            TabIndex        =   36
            Text            =   "FCompras.frx":171E
            ToolTipText     =   $"FCompras.frx":1723
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            TabIndex        =   38
            Text            =   " "
            Top             =   1080
            Width           =   1590
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
            Left            =   105
            TabIndex        =   91
            Top             =   1050
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
            Left            =   105
            TabIndex        =   90
            Top             =   735
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
            Left            =   105
            TabIndex        =   89
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
            TabIndex        =   88
            Top             =   210
            Width           =   1590
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
            Left            =   1470
            TabIndex        =   87
            Top             =   210
            Width           =   1590
         End
      End
      Begin VB.Frame FraDctoModificado 
         Caption         =   "NOTAS DE DEBITO/NOTAS DE CREDITO"
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
         Left            =   105
         TabIndex        =   39
         Top             =   4410
         Visible         =   0   'False
         Width           =   10095
         Begin VB.ComboBox CNumSerieTresComp 
            DataSource      =   "AdoAux"
            Height          =   315
            Left            =   6090
            TabIndex        =   43
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtNumSerieUnoComp 
            Height          =   330
            Left            =   5040
            MaxLength       =   3
            TabIndex        =   41
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumSerieDosComp 
            Height          =   336
            Left            =   5565
            MaxLength       =   3
            TabIndex        =   42
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumAutComp 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8715
            MaxLength       =   10
            TabIndex        =   45
            ToolTipText     =   $"FCompras.frx":17B9
            Top             =   432
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCDctoModif 
            Bindings        =   "FCompras.frx":1845
            DataSource      =   "AdoTipoComprobante"
            Height          =   288
            Left            =   108
            TabIndex        =   40
            ToolTipText     =   "Corresponde al tipo de comprobante que ha sido originalmente modificado antre la emisión de una nota de débito o crédito"
            Top             =   420
            Width           =   4848
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiComp 
            Height          =   330
            Left            =   7455
            TabIndex        =   44
            ToolTipText     =   $"FCompras.frx":1866
            Top             =   420
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin VB.Label Label29 
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
            Left            =   8715
            TabIndex        =   96
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha"
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
            Left            =   7455
            TabIndex        =   95
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie          Numero"
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
            Left            =   5040
            TabIndex        =   94
            Top             =   210
            Width           =   2325
         End
         Begin VB.Label Label26 
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
            TabIndex        =   93
            Top             =   210
            Width           =   4845
         End
      End
      Begin VB.Frame Frame1 
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
         Left            =   105
         TabIndex        =   15
         Top             =   1155
         Width           =   10095
         Begin VB.TextBox TxtBaseImpoNoObjIVA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3780
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   113
            Text            =   "FCompras.frx":1912
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtNumAutor 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            MaxLength       =   37
            TabIndex        =   20
            Text            =   "000000000000000000000000000000000001"
            Top             =   630
            Width           =   3300
         End
         Begin MSMask.MaskEdBox MBFechaCad 
            Height          =   330
            Left            =   2415
            TabIndex        =   23
            ToolTipText     =   $"FCompras.frx":1919
            Top             =   1365
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBFechaRegis 
            Height          =   330
            Left            =   1260
            TabIndex        =   22
            ToolTipText     =   $"FCompras.frx":19D0
            Top             =   1365
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   5250
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "FCompras.frx":1A58
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   5880
            MaxLength       =   9
            TabIndex        =   19
            Text            =   "0000001"
            ToolTipText     =   $"FCompras.frx":1A5F
            Top             =   630
            Width           =   855
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   5460
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   5040
            MaxLength       =   3
            TabIndex        =   17
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   630
            Width           =   435
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            MultiLine       =   -1  'True
            TabIndex        =   25
            Text            =   "FCompras.frx":1B02
            ToolTipText     =   $"FCompras.frx":1B09
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   8316
            MultiLine       =   -1  'True
            TabIndex        =   26
            Text            =   "FCompras.frx":1BB1
            ToolTipText     =   $"FCompras.frx":1BB6
            Top             =   1365
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCTipoComprobante 
            Bindings        =   "FCompras.frx":1C48
            DataSource      =   "AdoTipoComp"
            Height          =   315
            Left            =   105
            TabIndex        =   16
            ToolTipText     =   $"FCompras.frx":1C69
            Top             =   630
            Width           =   4950
            _ExtentX        =   8731
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
         Begin MSMask.MaskEdBox MBFechaEmi 
            Height          =   330
            Left            =   105
            TabIndex        =   21
            ToolTipText     =   $"FCompras.frx":1D11
            Top             =   1365
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.Label Label42 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NO OBJ. IVA"
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
            TabIndex        =   114
            Top             =   1050
            Width           =   1380
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
            Left            =   2415
            TabIndex        =   80
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label10 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " REGISTRO"
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
            TabIndex        =   79
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMISIÓN"
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
            TabIndex        =   78
            Top             =   1050
            Width           =   1170
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
            TabIndex        =   68
            Top             =   1050
            Width           =   1380
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
            TabIndex        =   82
            Top             =   1050
            Width           =   1380
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
            Left            =   5250
            TabIndex        =   81
            Top             =   1050
            Width           =   1380
         End
         Begin VB.Label Label3 
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
            Height          =   330
            Left            =   6720
            TabIndex        =   77
            Top             =   315
            Width           =   3270
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
            Height          =   330
            Left            =   5880
            TabIndex        =   76
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Height          =   330
            Left            =   5040
            TabIndex        =   75
            Top             =   315
            Width           =   855
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
            Height          =   330
            Left            =   105
            TabIndex        =   74
            Top             =   315
            Width           =   4950
         End
      End
      Begin VB.OptionButton OpcNo 
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
         Height          =   228
         Left            =   4305
         TabIndex        =   13
         ToolTipText     =   $"FCompras.frx":1DBD
         Top             =   315
         Value           =   -1  'True
         Width           =   636
      End
      Begin VB.OptionButton OpcSi 
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
         Left            =   2940
         TabIndex        =   12
         ToolTipText     =   $"FCompras.frx":1E55
         Top             =   315
         Width           =   636
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
         Height          =   4965
         Left            =   -74892
         TabIndex        =   47
         Top             =   315
         Width           =   10155
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FCompras.frx":1EED
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2415
            TabIndex        =   49
            Top             =   315
            Visible         =   0   'False
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            TabIndex        =   48
            Top             =   315
            Visible         =   0   'False
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8715
            TabIndex        =   58
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8085
            TabIndex        =   57
            Top             =   1470
            Width           =   645
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
            TabIndex        =   60
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   4515
            Width           =   1275
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
            Left            =   8085
            MultiLine       =   -1  'True
            TabIndex        =   54
            Text            =   "FCompras.frx":1F08
            Top             =   735
            Width           =   1905
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   56
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2415
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   53
            ToolTipText     =   $"FCompras.frx":1F0F
            Top             =   840
            Width           =   1590
         End
         Begin VB.TextBox TxtNumTresComRet 
            Height          =   336
            Left            =   1260
            MaxLength       =   9
            TabIndex        =   52
            Text            =   "0000001"
            ToolTipText     =   $"FCompras.frx":1F9B
            Top             =   840
            Width           =   1065
         End
         Begin VB.TextBox TxtNumDosComRet 
            Height          =   336
            Left            =   630
            MaxLength       =   3
            TabIndex        =   51
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   50
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FCompras.frx":203D
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   55
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
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FCompras.frx":205A
            Height          =   2595
            Left            =   105
            TabIndex        =   59
            Top             =   1890
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   4577
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
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
         Begin VB.Label Label41 
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
            TabIndex        =   108
            Top             =   4515
            Width           =   2010
         End
         Begin VB.Label Label40 
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
            TabIndex        =   107
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label36 
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
            TabIndex        =   106
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label35 
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
            TabIndex        =   105
            Top             =   1260
            Width           =   1380
         End
         Begin VB.Label Label34 
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
            TabIndex        =   104
            Top             =   1260
            Width           =   6630
         End
         Begin VB.Label Label33 
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
            Left            =   6720
            TabIndex        =   103
            Top             =   735
            Width           =   1380
         End
         Begin VB.Label Label32 
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
            TabIndex        =   102
            Top             =   630
            Width           =   1590
         End
         Begin VB.Label Label31 
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
            TabIndex        =   101
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label30 
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
            TabIndex        =   100
            Top             =   630
            Width           =   1065
         End
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FCompras.frx":2076
         DataSource      =   "AdoSustento"
         Height          =   315
         Left            =   105
         TabIndex        =   14
         ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
         Top             =   840
         Width           =   9465
         _ExtentX        =   16695
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
      Begin VB.PictureBox Label23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         ScaleHeight     =   165
         ScaleWidth      =   1845
         TabIndex        =   61
         Top             =   2310
         Width           =   1905
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DE SUSTENTO TRIBUTARIO"
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
         Top             =   630
         Width           =   9465
      End
      Begin VB.Label Label15 
         Caption         =   " DEVOLUCION DE I.V.A. "
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
         TabIndex        =   69
         Top             =   315
         Width           =   2535
      End
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2730
      Top             =   3045
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoTipoIdentificacion 
      Height          =   330
      Left            =   210
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   2415
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   210
      Top             =   2730
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicios 
      Height          =   330
      Left            =   210
      Top             =   3045
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoPorIva 
      Height          =   330
      Left            =   2730
      Top             =   2730
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoPorIce 
      Height          =   330
      Left            =   2730
      Top             =   2415
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoCaTrTiCom 
      Height          =   330
      Left            =   210
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Catal.Tributarios y Tipos de Comprobantes"
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
      Top             =   3675
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoConceptoRet"
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
   Begin MSAdodcLib.Adodc AdoTransCompras 
      Height          =   330
      Left            =   210
      Top             =   3990
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   210
      Top             =   4305
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AsientoAir"
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
      Left            =   2730
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   2730
      Top             =   3360
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteServicios"
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
      Left            =   2730
      Top             =   3675
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteBienes"
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
   Begin MSAdodcLib.Adodc AdoAsientoCompras 
      Height          =   330
      Left            =   210
      Top             =   4935
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   210
      Top             =   5250
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DCProveedor 
      Bindings        =   "FCompras.frx":2090
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   105
      TabIndex        =   9
      ToolTipText     =   "En este combo de selección se despliega una lista de todos los proveedores"
      Top             =   1575
      Width           =   9150
      _ExtentX        =   16140
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
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   210
      Top             =   4620
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransAir"
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
      Left            =   2730
      Top             =   3990
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
      Left            =   105
      TabIndex        =   92
      Top             =   1365
      Width           =   9150
   End
   Begin VB.Label LblMensaje 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   105
      TabIndex        =   70
      Top             =   105
      Width           =   9255
   End
   Begin VB.Label LblNumIdent 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   7560
      TabIndex        =   11
      Top             =   1575
      Width           =   1800
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
      TabIndex        =   10
      Top             =   1575
      Width           =   330
   End
End
Attribute VB_Name = "FCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim FechaRegis As Date
Dim CapDm, CapDcto, Cap1, Captc, Cap, Ct, Espizq, Espder, ValorP, ch, Ch1, CodProv, CodSus As String
Dim OP As Boolean
Dim Rb, Rs, Rf, cod As Byte
Dim AniocadAux, SumAnio, Aniocad As Integer
Dim CalIsMi, CalmIva, CalIbMi, ac As Double

Private Sub ChRetB_Click()
  If ChRetB.value <> 0 Then
     ch = 1
     Ch1 = "B"
     DCRetIBienes.Visible = True
     TxtIvaBienMonIva.Enabled = True
     DCPorcenRetenIvaBien.Enabled = True
     TxtIvaBienValRet.Enabled = True
  Else
     Ch1 = "S"
     DCRetIBienes.Visible = False
     TxtIvaBienMonIva.Enabled = False
     DCPorcenRetenIvaBien.Enabled = False
     TxtIvaBienValRet.Enabled = False
  End If
  If ChRetB.value <> 0 And ChRetB.value <> 0 Then
     Ch1 = "X"
  End If
End Sub

Private Sub ChRetS_Click()
  If ChRetS.value <> 0 Then
     ch = 1
     Ch1 = "S"
     'Carga_ConceptosVerif
     DCRetISer.Visible = True
     TxtIvaSerMonIva.Enabled = True
     DCPorcenRetenIvaServ.Enabled = True
     TxtIvaSerValRet.Enabled = True
  Else
     DCRetISer.Visible = False
     TxtIvaSerMonIva.Enabled = False
     DCPorcenRetenIvaServ.Enabled = False
     TxtIvaSerValRet.Enabled = False
  End If
  If ChRetB.value <> 0 And ChRetS.value <> 0 Then
     Ch1 = "X"
     'Carga_ConceptosVerif
  End If
End Sub

Private Sub CmdAir_Click()
  SSTCompras.Tab = 1
  TxtNumUnoComRet.SetFocus
End Sub

Private Sub CmdCerrar_Click()
  'Borra Asiento Compras
  sSQL = "DELETE * " _
       & "FROM Asiento_Compras " _
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
       & "AND Tipo_Trans = 'C' "
  ConectarAdoExecute sSQL
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim ProvAnt As String
  ProvAnt = DCProveedor
 'Valido por si acaso exista algun valor con 0
  TextoValido TxtIvaBienMonIva, True, , 2
  TextoValido TxtBaseImpo, True, , 2
  TextoValido TxtBaseImpoGrav, True, , 2
  TextoValido TxtBaseImpoIce, True, , 2
  TextoValido TxtMontoIva, True, , 2
  TextoValido TxtMontoIce, True, , 2
  TextoValido TxtIvaBienMonIva, True, , 2
  TextoValido TxtIvaBienValRet, True, , 2
  TextoValido TxtIvaSerMonIva, True, , 2
  TextoValido TxtIvaSerValRet, True, , 2
  'Pregunto antes de grabar
  Titulo = "GRABAR COMPRAS"
  Mensajes = "Desea Grabar los Datos"
  If BoxMensaje = vbYes Then
    'Borrar todas las transacciones de compras que tengan la misma factura y la misma retencion
    'del mismo proveedor
     Eliminar_Trans_AT "C", CodigoCliente, TxtNumSerieUno, TxtNumSerieDos, TxtNumSerietres, TxtNumAutor, "", "", "", CMes, CAño, Ln_SRI
     Eliminar_Trans_Air "C", CodigoCliente, CMes, CAño, Ln_SRI
     Grabacion
     If Ln_SRI < 0 Then
        Mensajes = "Los Datos fueron grabados correctamente" & vbCrLf _
                 & "Desea ingresar otra transacción"""
        If BoxMensaje = vbYes Then
           Ln_No = 1
           Limpiar_Controles
           Listar_Air
           SSTCompras.Tab = 0
           DCProveedor = ProvAnt
           CTP.SetFocus
        Else
           Unload FCompras
        End If
     Else
        MsgBox "Datos Modificados Correctamente"
        Unload FCompras
     End If
   Else
      ChRetB.SetFocus
   End If
End Sub

Private Sub CMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMes_LostFocus()
  Habilita_Controles
  Modificacion_AT "C", CModificacion, CMes, CAño
End Sub

Private Sub CModificacion_GotFocus()
  MarcarTexto CModificacion
End Sub

Private Sub CModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CModificacion_LostFocus()
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
      & "FROM Trans_Compras " _
      & "WHERE Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
      & "AND Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND Linea_SRI = " & I & " " _
      & "AND IdProv = '" & CodigoCliente & "' " _
      & "ORDER BY Linea_SRI "
 SelectAdodc AdoTransCompras, sSQL
 With AdoTransCompras.Recordset
  If .RecordCount > 0 Then
     'Busco el Proveedor
      Ln_SRI = I
      If .Fields("Cta_Bienes") <> Ninguno Then
          DCRetIBienes = .Fields("Cta_Bienes")
          ChRetB.value = 1
      End If
      If .Fields("Cta_Servicio") <> Ninguno Then
          DCRetISer = .Fields("Cta_Servicio")
          ChRetS.value = 1
      End If
      CodProv = .Fields("IdProv")
      If AdoClientes.Recordset.RecordCount > 0 Then
         AdoClientes.Recordset.MoveFirst
         AdoClientes.Recordset.Find ("Codigo = '" & CodProv & "' ")
         If Not AdoClientes.Recordset.EOF Then
            DCProveedor = AdoClientes.Recordset.Fields("Cliente")
            LblTD = AdoClientes.Recordset.Fields("TD")
            LblNumIdent = AdoClientes.Recordset.Fields("CI_RUC")
            CodigoCliente = AdoClientes.Recordset.Fields("Codigo")
            DireccionCli = AdoClientes.Recordset.Fields("Direccion")
            CICliente = LblNumIdent
            TipoBenef = LblTD
           'Si existe beneficiario
            If .Fields("DevIva") = "S" Then OpcSi.value = True Else OpcNo.value = True
           'Busco el Sustento Tributario para capturar la descripcion
            CodSus = .Fields("CodSustento")
            If AdoSustento.Recordset.RecordCount > 0 Then
               AdoSustento.Recordset.MoveFirst
               AdoSustento.Recordset.Find ("Credito_Tributario = '" & CodSus & "' ")
               If Not AdoSustento.Recordset.EOF Then
                  DCSustento = CodSus & " - " & AdoSustento.Recordset.Fields("Descripcion")
               Else
                  MsgBox "Este Código no existe", vbInformation, "Aviso"
               End If
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
                  DCTipoComprobante = AdoTipoComprobante.Recordset.Fields("Descripcion")
               Else
                  MsgBox "El Comprobante no existe", vbInformation, "Aviso"
               End If
            End If
            MBFechaRegis = .Fields("FechaRegistro")
           'Carga_PorcentajeIva (MBFechaRegis)
           'Es un numero el campo PorcentajeIva y lo convierto a String
            CodPorIva = CStr(.Fields("PorcentajeIva"))
            If AdoPorIva.Recordset.RecordCount > 0 Then
               AdoPorIva.Recordset.MoveFirst
               AdoPorIva.Recordset.Find ("Codigo = '" & CodPorIva & "' ")
               If Not AdoPorIva.Recordset.EOF Then
                  DCPorcenIva = AdoPorIva.Recordset.Fields("Porc")
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
                  DCPorcenIce = AdoPorIce.Recordset.Fields("Porc")
               Else
                  MsgBox "Porcentaje de ICE no existe", vbInformation, "Aviso"
               End If
            End If
           'Cargo el Codigo retencion IVA Bienes
            CodRetBien = .Fields("PorRetBienes")
            If AdoRetIvaBienes.Recordset.RecordCount > 0 Then
               AdoRetIvaBienes.Recordset.MoveFirst
               AdoRetIvaBienes.Recordset.Find ("Codigo = '" & CodRetBien & "' ")
               If Not AdoRetIvaBienes.Recordset.EOF Then
                  DCPorcenRetenIvaBien = AdoRetIvaBienes.Recordset.Fields("Porc")
               Else
                  MsgBox "Código de Retención no existe", vbInformation, "Aviso"
               End If
            End If
           'Cargo el Codigo retencion IVA Servicios
            CodRetServ = .Fields("PorRetServicios")
            If AdoRetIvaServicios.Recordset.RecordCount > 0 Then
               AdoRetIvaServicios.Recordset.MoveFirst
               AdoRetIvaServicios.Recordset.Find ("Codigo = '" & CodRetServ & "' ")
               If Not AdoRetIvaServicios.Recordset.EOF Then
                  DCPorcenRetenIvaServ = AdoRetIvaServicios.Recordset.Fields("Porc")
               Else
                  MsgBox "Código de Retención no existe", vbInformation, "Aviso"
               End If
            End If
            CTP.AddItem .Fields("TP")
            TxtNumSerieUno = .Fields("Establecimiento")
            TxtNumSerieDos = .Fields("PuntoEmision")
            TxtNumSerietres = .Fields("Secuencial")
            TxtNumAutor = .Fields("Autorizacion")
            TxtNumeroC = .Fields("Numero")
            MBFechaEmi = .Fields("FechaEmision")
            MBFechaCad = .Fields("FechaCaducidad")
            TxtBaseImpo = .Fields("BaseImponible")
            TxtBaseImpoGrav = .Fields("BaseImpGrav")
            TxtMontoIva = .Fields("MontoIva")
            TxtBaseImpoIce = .Fields("BaseImpIce")
            TxtMontoIce = .Fields("MontoIce")
            TxtIvaBienMonIva = .Fields("MontoIvaBienes")
            TxtIvaBienValRet = .Fields("ValorRetBienes")
            TxtIvaSerMonIva = .Fields("MontoIvaServicios")
            TxtIvaSerValRet = .Fields("ValorRetServicios")
            TxtNumConParPol = .Fields("ContratoPartidoPolitico")
            TxtMonTitOner = .Fields("MontoTituloOneroso")
            TxtMonTitGrat = .Fields("MontoTituloGratuito")
            
           'Busco en el Trans Air para ver si tiene retenciones
            Mifecha = BuscarFecha(MBFechaRegis)
            sSQL = "DELETE * " _
                 & "FROM Asiento_Air " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND CodigoU = '" & CodigoUsuario & "' " _
                 & "AND T_No = " & Trans_No & " " _
                 & "AND Tipo_Trans = 'C' "
            ConectarAdoExecute sSQL
            Ln_No = Maximo_De("Asiento_Air", "A_No")
            sSQL = "SELECT TA.*, TC.Concepto " _
                 & "FROM Trans_Air As TA, Tipo_Concepto_Retencion As TC " _
                 & "WHERE TA.Fecha Between #" & FechaIni & "# AND #" & FechaFin & "# " _
                 & "AND TA.Item = '" & NumEmpresa & "' " _
                 & "AND TA.Periodo = '" & Periodo_Contable & "' " _
                 & "AND TA.IdProv = '" & CodProv & "' " _
                 & "AND TC.Fecha_Inicio <= #" & Mifecha & "# " _
                 & "AND TC.Fecha_Final >= #" & Mifecha & "# " _
                 & "AND Tipo_Trans = 'C' " _
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
                  SetAdoFields "Tipo_Trans", "C"
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
                 & "AND Tipo_Trans = 'C' " _
                 & "ORDER BY CodRet "
            SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
           'Pongo la Base Imponible
            TxtSumatoria = CTNumero(TxtBaseImpoNoObjIVA, 2) + CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2)
         Else
            MsgBox "Este beneficiario no existe", vbInformation, "Aviso"
         End If
      End If
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

Private Sub DCConceptoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConceptoRet_LostFocus()
  If IsNumeric(DCConceptoRet.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCConceptoRet.SetFocus
  Else
     With AdoConceptoRet.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          OP = False
         .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRet) & "' ")
          Cadena = .Fields("Ingresar_Porcentaje")
          TxtPorRetConA.Enabled = False
          If Not .EOF Then
            'Verifico el codigo para activar el text e ingrese el porcentaje manualmente
            Select Case Cadena
               Case "N": TxtPorRetConA = .Fields("Porcentaje")
                    OP = True
               Case "S": TxtPorRetConA.Enabled = True
            End Select
          Else
             MsgBox "No encontro este código vuelva a buscar"
          End If
      End If
     End With
     TxtBimpConA = TxtSumatoria
  End If
End Sub

Private Sub DCDctoModif_LostFocus()
  Captura_TipoComprobante_DctoModificado
End Sub

Private Sub DCPorcenIva_GotFocus()
  MarcarTexto DCPorcenIva
End Sub

Private Sub DCPorcenIva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProveedor_LostFocus()
  SSTCompras.Tab = 0
  If IsNumeric(DCProveedor.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCProveedor.Text = ""
     Leer_Clientes
     DCProveedor.SetFocus
  Else
     NombreCliente = UCase$(DCProveedor)
     With AdoClientes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & NombreCliente & "' ")
          If Not .EOF Then
             Label41.Caption = "PROVEEDOR: "
            'Busca y captura el codigo de Porcentaje IVA
            CodigoCliente = .Fields("Codigo")
            DireccionCli = .Fields("Direccion")
            CICliente = .Fields("CI_RUC")
            TipoBenef = .Fields("TD")
            If .Fields("RISE") Then Label41.Caption = Label41.Caption & " RISE"
            If .Fields("Especial") Then Label41.Caption = Label41.Caption & "Contribuyente especial "
            LblNumIdent = CICliente
            LblTD.Caption = TipoBenef
              
            TxtNumSerietres = "0000001"
            'Aqui despliego el ultimo numero de la Transaccion
            sSQL = "SELECT TOP 1 * " _
                 & "FROM Trans_Compras " _
                 & "WHERE IdProv = '" & CodigoCliente & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Secuencial DESC "
            SelectAdodc AdoAux, sSQL
            With AdoAux.Recordset
             If .RecordCount > 0 Then TxtNumSerietres = .Fields("Secuencial")
            End With
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

Private Sub DCTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_LostFocus()
  If IsNumeric(DCTipoComprobante.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCTipoComprobante.Text = ""
     Carga_TipoComprobante (DCSustento)
     DCTipoComprobante.SetFocus
     Captura_TipoComprobante
  Else
     Captura_TipoComprobante
  End If
End Sub

Private Sub DGConceptoAir_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     Titulo = "Aviso"
     Mensajes = "Desea Eliminar la Retención"
     If BoxMensaje = vbYes Then
        With AdoAsientoAir.Recordset
         If .RecordCount > 0 Then
             Codigo = .Fields("CodRet")
             No_Desde = .Fields("SecRetencion")
             Mifecha = BuscarFecha(.Fields("FechaEmiRet"))
             Codigo1 = .Fields("AutRetencion")
             J = .Fields("A_No")
             sSQL = "DELETE * " _
                  & "FROM Asiento_Air " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND IdProv = '" & CodigoCliente & "' " _
                  & "AND T_No = " & Trans_No & " " _
                  & "AND Tipo_Trans = 'C' " _
                  & "AND A_No = " & J & " " _
                  & "AND CodRet = '" & Codigo & "' "
             ConectarAdoExecute sSQL
         End If
         sSQL = "SELECT * " _
              & "FROM Asiento_Air " _
              & "WHERE CodRet <> '.' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' " _
              & "AND T_No = " & Trans_No & " " _
              & "AND Tipo_Trans = 'C' " _
              & "ORDER BY CodRet "
         SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
         Calculo_Sumatoria
        End With
  End If
End If
End Sub

Private Sub MBFechaCad_GotFocus()
  MarcarTexto MBFechaCad
End Sub

Private Sub MBFechaCad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmi_GotFocus()
  MarcarTexto MBFechaEmi
End Sub

Private Sub MBFechaEmi_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_GotFocus()
  MarcarTexto MBFechaEmiComp
End Sub

Private Sub MBFechaEmiComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_LostFocus()
  FechaValida MBFechaEmiComp
End Sub

Private Sub OpcNo_LostFocus()
  If OpcNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcSi_LostFocus()
  If OpcSi.value = True Then ValorP = "S"
End Sub

Private Sub SSTCompras_Click(PreviousTab As Integer)
  Select Case PreviousTab
         Case 0: If ChRetF.Visible Then ChRetF.SetFocus Else CTP.SetFocus
         Case 1: OpcSi.SetFocus
  End Select
End Sub

Private Sub TxtBaseImpo_GotFocus()
  MarcarTexto TxtBaseImpo
End Sub

Private Sub TxtBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpo_LostFocus()
  TextoValido TxtBaseImpo, True, , 0
End Sub

Private Sub TxtBaseImpoGrav_GotFocus()
  MarcarTexto TxtBaseImpoGrav
End Sub

Private Sub TxtBaseImpoGrav_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoGrav_LostFocus()
  TextoValido TxtBaseImpoGrav, True, , 0
End Sub

Private Sub TxtBaseImpoIce_GotFocus()
  MarcarTexto TxtBaseImpoIce
End Sub

Private Sub TxtBaseImpoIce_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIce_LostFocus()
  TextoValido TxtBaseImpoIce, True, , 0
End Sub

Private Sub TxtBimpConA_GotFocus()
  MarcarTexto TxtBimpConA
End Sub

Private Sub TxtBimpConA_LostFocus()
  RatonNormal
  TextoValido TxtBimpConA, True, , 0
  TextoValido TxtSumatoria, True, , 0
  'Valida que la base imponible no sea mayor que la BIG y la BIcero
  If CTNumero(TxtBimpConA, 2) > CTNumero(TxtSumatoria, 2) Then
     MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
     & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
     TxtBimpConA.Text = 0
     TxtBimpConA.SetFocus
  Else
     If OP Then
        TxtValConA = (CTNumero(TxtBimpConA, 2) * CTNumero(TxtPorRetConA, 2)) / 100
        Insertar_AsientoAir
        If (cod = 4) Or (cod = 5) Then
           DCDctoModif.SetFocus
        Else
           TxtNumConParPol.SetFocus
        End If
     Else
        'TxtPorRetConA.SetFocus
     End If
  End If
End Sub

Sub Insertar_AsientoAir()
'Selecciona el numero mayor para continuar la secuencia en el
'campo T_No y A_No
Ln_No = Maximo_De("Asiento_Air", "A_No")
If CTNumero(TxtBimpConA, 2) > 0 Then
   RatonReloj
   Espizq = SinEspaciosIzq(DCConceptoRet)
   Espder = Trim$(Mid$(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
   SetAdoAddNew "Asiento_Air"
   SetAdoFields "CodRet", Espizq
   SetAdoFields "Detalle", Espder
   SetAdoFields "BaseImp", CTNumero(TxtBimpConA, 2)
   SetAdoFields "Porcentaje", CTNumero(TxtPorRetConA, 2) / 100
   SetAdoFields "ValRet", CTNumero(TxtValConA, 2)
   SetAdoFields "EstabRetencion", TxtNumUnoComRet
   SetAdoFields "PtoEmiRetencion", TxtNumDosComRet
   SetAdoFields "SecRetencion", CTNumero(TxtNumTresComRet)
   SetAdoFields "AutRetencion", TxtNumUnoAutComRet
   SetAdoFields "FechaEmiRet", MBFechaRegis
   SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
   SetAdoFields "EstabFactura", TxtNumSerieUno
   SetAdoFields "PuntoEmiFactura", TxtNumSerieDos
   SetAdoFields "Factura_No", CTNumero(TxtNumSerietres)
   SetAdoFields "IdProv", CodigoCliente
   SetAdoFields "A_No", Ln_No
   SetAdoFields "T_No", Trans_No
   SetAdoFields "Tipo_Trans", "C"
   SetAdoUpdate
          
  'Despliega los datos en el DataGrid
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE CodRet <> '.' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
   
  'Se situa en el combo de retención AIR
   If ChRetF.Visible Then DCRetFuente.SetFocus Else TxtNumUnoComRet.SetFocus
   
  'Realiza la Sumatoria de las Retenciones
   ac = ac + TxtValConA
   TxtTotalReten = ac
End If
RatonNormal
End Sub

Private Sub DCPorcenIce_LostFocus()
  If Not IsNumeric(DCPorcenIce) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIce = ""
    'Carga_PorcentajeIce
     DCPorcenIce.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIce = "0"
     With AdoPorIce.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CSng(DCPorcenIce) & " ")
          If Not .EOF Then CodPorIce = .Fields("Codigo")
       End If
      End With
      Total_IVA = 0
      Total_IVA = CTNumero(TxtBaseImpoIce, 2)
      TxtMontoIce = 0
     'Calcula el Porcentaje de Ice
      CalIbMi = Total_IVA * CTNumero(DCPorcenIce, 2) / 100
      TxtMontoIce = CalIbMi
  End If

 'Coloca el valor de Monto IVA dependiendo si se activo Bienes o Servicios
  If ChRetB + ChRetS = 0 Then
     TxtIvaBienMonIva = TxtMontoIva
  End If
  If ChRetB.value <> 0 Then
     TxtIvaBienMonIva = TxtMontoIva
     TxtIvaSerMonIva = 0
  Else
     If ChRetS.value <> 0 Then
        TxtIvaSerMonIva = TxtMontoIva
        TxtIvaBienMonIva = 0
     End If
  End If
End Sub

Private Sub DCPorcenIva_LostFocus()
  If Not IsNumeric(DCPorcenIva) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenIva = ""
    'Carga_PorcentajeIva (MBFechaRegis)
     DCPorcenIva.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje IVA
     CodPorIva = "0"
     With AdoPorIva.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & CByte(DCPorcenIva) & " ")
          If Not .EOF Then CodPorIva = .Fields("Codigo")
      End If
     End With
     Total_IVA = 0
     Total_IVA = CTNumero(TxtBaseImpoGrav, 2)
    'Calcula el Porcentaje de Iva
     CalmIva = (Total_IVA * DCPorcenIva) / 100
     TxtMontoIva = CalmIva
  End If
End Sub

Private Sub DCPorcenRetenIvaBien_LostFocus()
  CodRetBien = 0
  If Not IsNumeric(DCPorcenRetenIvaBien) Then
     MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCPorcenRetenIvaBien = ""
     Carga_RetencionIvaBienes_Servicios
     DCPorcenRetenIvaBien.SetFocus
  Else
    'Busca y captura el codigo de Porcentaje retencion Iva Bienes
     With AdoRetIvaBienes.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBien) & " ")
          If Not .EOF Then CodRetBien = .Fields("Codigo")
      Else
         MsgBox "Código incorrecto", vbInformation, "Aviso"
      End If
     End With
     Total_IVA = CTNumero(TxtIvaBienMonIva, 2)
    'Calcula la retencion Iva Bienes
     CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBien)) / 100
     TxtIvaBienValRet = CalIbMi
  End If
  TxtIvaSerMonIva = Format(CTNumero(TxtMontoIva, 2) - CTNumero(TxtIvaBienMonIva, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServ_LostFocus()
  CodRetServ = 0
 'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
  If DCPorcenRetenIvaServ = "70/100" Then
     Ct = "Si"
     TxtIvaSerValRet.Text = ""
     TxtIvaSerValRet.Enabled = True
     TxtIvaSerValRet.SetFocus
  Else
     If Not IsNumeric(DCPorcenRetenIvaServ) Then
        MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
        DCPorcenRetenIvaServ = ""
        Carga_RetencionIvaBienes_Servicios
        DCPorcenRetenIvaServ.SetFocus
     End If
  End If
    
 'Busca captura el codigo de Porcentaje retencion Iva Servicios
  With AdoRetIvaServicios.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServ) & " ")
       If Not .EOF Then CodRetServ = .Fields("Codigo")
   Else
      MsgBox "Código Incorrecto", vbInformation, "Aviso"
   End If
  End With
  Ct = "No"
  Total_IVA = 0
  Total_IVA = CTNumero(TxtIvaSerMonIva, 2)
  If DCPorcenRetenIvaServ = "70/100" Then
  Else
     CalIsMi = (Total_IVA * CInt(DCPorcenRetenIvaServ)) / 100
     TxtIvaSerValRet = CalIsMi
     TxtIvaSerValRet.Enabled = False
  End If
  SSTCompras.Tab = 0
  SSTCompras.SetFocus
End Sub

Private Sub DCSustento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSustento_LostFocus()
  If IsNumeric(DCSustento.Text) Then
     MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
     DCSustento.Text = ""
     Carga_CreditoTributario
     DCSustento.SetFocus
     Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
  Else
     Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
  End If
End Sub

Private Sub Form_Activate()
  Ln_No = 1
  Ln_SRI = -1
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
  CentrarForm FCompras
  ConectarAdodc AdoAux
  ConectarAdodc AdoSustento
  ConectarAdodc AdoTipoIdentificacion
  ConectarAdodc AdoTipoComprobante
  ConectarAdodc AdoRetIvaBienes
  ConectarAdodc AdoRetIvaServicios
  ConectarAdodc AdoPorIce
  ConectarAdodc AdoPorIva
  ConectarAdodc AdoConceptoRet
  ConectarAdodc AdoAsientoCompras
  ConectarAdodc AdoTransCompras
  ConectarAdodc AdoAsientoAir
  ConectarAdodc AdoTransAir
  ConectarAdodc AdoClientes
  ConectarAdodc AdoRetFuente
  ConectarAdodc AdoRetIvaSerCC
  ConectarAdodc AdoRetIvaBienesCC
End Sub

Private Sub MBFechaCad_LostFocus()
  'Verifico que la fecha de caducidad no sea mayor a la de emisión
   FechaValida MBFechaCad
   If MBFechaCad = "00/00/0000" Then
      MsgBox "Fecha no válida, vuelva a ingresar", vbInformation, "Aviso"
      MBFechaCad.SetFocus
   Else
        'Captura el año de la fecha de emisión
        Anio = Year(MBFechaEmi)
        SumAnio = Anio + 1  'Emisión + 1 año
        Aniocad = Year(MBFechaCad)
        AniocadAux = Aniocad + 1 'Asigno en otra variable el año de caducidad
        'Verifica si el año de caducidad es menor que el año de Emisión
        If (Aniocad < Anio) Then
           MsgBox "La Fecha de Caducidad no debe ser < a la Fecha de Emisión", vbInformation, "Aviso"
           FechaValida MBFechaCad
           MBFechaCad.SetFocus
        Else
           'Verifica si el año de caducidad es mayor con 2 años al año de Emisión
           If (Aniocad = AniocadAux) Then
              MsgBox "Hola La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
              FechaValida MBFechaCad
              MBFechaCad.SetFocus
           Else
             If (Aniocad > AniocadAux) Then
                MsgBox "La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
                FechaValida MBFechaCad
                MBFechaCad.SetFocus
             End If
           End If
        End If
 End If
End Sub

Private Sub MBFechaEmi_LostFocus()
  FechaValida MBFechaEmi
 'Controla que la Fecha de Emisiòn este entre 01/01/2000 en adelante
  If CFechaLong(MBFechaEmi) < CFechaLong("01/01/2000") Then
     MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
     MBFechaEmi = "01/01/2000"
     MBFechaEmi.SetFocus
  End If
  MBFechaRegis = MBFechaEmi
End Sub

Private Sub MBFechaRegis_GotFocus()
  MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegis_LostFocus()
  FechaValida MBFechaRegis
  'Controla que la Fecha de Registro este entre 01/01/2000 en adelante
  If CFechaLong(MBFechaRegis) < CFechaLong("01/01/2000") Then
     MsgBox "La Fecha de Registro debe ser mayor que 01/01/2000", vbInformation, "Aviso"
     MBFechaRegis = "01/01/2000"
     MBFechaRegis.SetFocus
  Else
     If MBFechaRegis < MBFechaEmi Then
        MsgBox "La Fecha de Registro debe ser mayor o igual que la Fecha de Emisión", vbInformation, "Aviso"
        MBFechaRegis.SetFocus
     End If
  End If
  FechaValida MBFechaRegis
 'Carga la Tabla de Porcentaje Iva en el DataCombo
 'Carga_PorcentajeIva (MBFechaRegis)
  Carga_ConceptosRetencion MBFechaRegis
End Sub

Private Sub TxtIvaBienMonIva_GotFocus()
  MarcarTexto TxtIvaBienMonIva
End Sub

Private Sub TxtIvaBienMonIva_LostFocus()
  'MsgBox CTNumero(TxtIvaBienMonIva, 2)
  TextoValido TxtIvaBienMonIva, True, , 0
End Sub


Private Sub TxtIvaSerMonIva_GotFocus()
  MarcarTexto TxtIvaSerMonIva
End Sub

Private Sub TxtIvaSerMonIva_LostFocus()
  TextoValido TxtIvaSerMonIva, True, , 0
  'Verifica el Monto Iva Servicios
  If CDbl(TxtIvaBienMonIva) + CDbl(TxtIvaSerMonIva) > CDbl(TxtMontoIva) Then
     MsgBox "Monto IVA Servicios no puede ser > que Monto IVA", vbInformation, "Aviso de Compras"
     TxtIvaSerMonIva.Text = ""
     TxtIvaSerMonIva.SetFocus
  End If
End Sub

Private Sub TxtIvaSerValRet_GotFocus()
  MarcarTexto TxtIvaSerValRet
End Sub

Private Sub TxtIvaSerValRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitGrat_GotFocus()
  MarcarTexto TxtMonTitGrat
End Sub

Private Sub TxtMonTitGrat_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitGrat_LostFocus()
  TextoValido TxtMonTitGrat, True, , 2
End Sub

Private Sub TxtMonTitOner_GotFocus()
  MarcarTexto TxtMonTitOner
End Sub

Private Sub TxtMonTitOner_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitOner_LostFocus()
  TextoValido TxtMonTitOner, True, , 0
End Sub

Private Sub TxtMontoIva_GotFocus()
  MarcarTexto TxtMontoIva
End Sub

Private Sub TxtMontoIva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIva_LostFocus()
  TextoValido TxtMontoIva, True, , 0
End Sub

Private Sub TxtNumAutComp_GotFocus()
  MarcarTexto TxtNumAutComp
End Sub

Private Sub TxtNumAutComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutComp_LostFocus()
  If Val(TxtNumAutComp) <= 0 Then TxtNumAutComp = "0000000000"
  TxtNumAutComp = Format(Val(Round(TxtNumAutComp)), String(10, "0"))
End Sub

Private Sub TxtNumAutor_GotFocus()
   MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
  If Val(TxtNumAutor) <= 0 Then TxtNumAutor = "0000000001"
End Sub

Private Sub TxtNumConParPol_GotFocus()
  MarcarTexto TxtNumConParPol
End Sub

Private Sub TxtNumConParPol_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumConParPol_LostFocus()
  TextoValido TxtNumConParPol, True, , 0
  TxtNumConParPol = Format(Val(CCur(TxtNumConParPol)), String(10, "0"))
End Sub

Private Sub TxtNumDosComRet_GotFocus()
  MarcarTexto TxtNumDosComRet
End Sub

Private Sub TxtNumDosComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumDosComRet_LostFocus()
  TextoValido TxtNumDosComRet, True, , 0
  If Val(TxtNumDosComRet) <= 0 Then TxtNumDosComRet = "001"
  TxtNumDosComRet = Format(Val(TxtNumDosComRet), "000")
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

Private Sub TxtNumSerieDosComp_GotFocus()
  MarcarTexto TxtNumSerieDosComp
End Sub

Private Sub TxtNumSerieDosComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDosComp_LostFocus()
  TextoValido TxtNumSerieDosComp, True, , 0
  If Val(TxtNumSerieDosComp) <= 0 Then TxtNumSerieDosComp = "001"
  TxtNumSerieDosComp = Format(Val(TxtNumSerieDosComp), "000")
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

Private Sub TxtNumSerieUnoComp_GotFocus()
  MarcarTexto TxtNumSerieUnoComp
End Sub

Private Sub TxtNumSerieUnoComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUnoComp_LostFocus()
  TextoValido TxtNumSerieUnoComp, True, , 0
  If Val(TxtNumSerieUnoComp) <= 0 Then TxtNumSerieUnoComp = "001"
  TxtNumSerieUnoComp = Format(Val(TxtNumSerieUnoComp), "000")
End Sub

Private Sub TxtNumTresComRet_GotFocus()
  MarcarTexto TxtNumTresComRet
End Sub

Private Sub TxtNumTresComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRet_LostFocus()
  If Val(TxtNumTresComRet) <= 0 Then TxtNumTresComRet = "000000000"
  TxtNumTresComRet = Format(Val(Round(TxtNumTresComRet)), "000000000")
  'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
  TxtSumatoria = CTNumero(TxtBaseImpoNoObjIVA, 2) + CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2)
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
  MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRet_LostFocus()
  If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0000000000"
  TxtNumUnoAutComRet = Format(Val(Round(TxtNumUnoAutComRet)), String(10, "0"))
End Sub

Private Sub TxtNumUnoComRet_GotFocus()
  MarcarTexto TxtNumUnoComRet
End Sub

Private Sub TxtNumUnoComRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRet_LostFocus()
  TextoValido TxtNumUnoComRet, True, , 0
  If Val(TxtNumUnoComRet) <= 0 Then TxtNumUnoComRet = "001"
  TxtNumUnoComRet = Format(Val(TxtNumUnoComRet), "000")
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND TD <>  'E' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCProveedor, AdoClientes, sSQL, "Cliente"
End Sub

Public Sub Carga_CreditoTributario()
  'Carga la Tabla de Catalogos Tributarios al DataCombo
  sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento,* " _
       & "FROM Tipo_Tributario " _
       & "WHERE Credito_Tributario <> '.' " _
       & "ORDER BY Credito_Tributario "
  SelectDBCombo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoComprobante(CargaTC As String)
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo <> 100 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    
    'Capturo el codigo del Tipo de Catalogo Tributario
     Cap = CargaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
    Cadena = Ninguno
    With AdoSustento.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Credito_Tributario = '" & CargaTC & "' ")
         If Not .EOF Then Cadena = .Fields("Codigo_Tipo_Comprobante")
     End If
    End With
    sSQL = "SELECT * " _
         & "FROM Tipo_Comprobante " _
         & "WHERE Tipo_Comprobante_Codigo IN (" & Cadena & ") "
    If TipoBenef = "R" Then
       sSQL = sSQL & "AND R <> " & Val(adFalse) & " "
    Else
       sSQL = sSQL & "AND C <> " & Val(adFalse) & " "
    End If
    sSQL = sSQL & "ORDER BY Tipo_Comprobante_Codigo "
    SelectDBCombo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Captura_TipoComprobante()
  'Captura lo que tiene el Combo de Tipo de Comprobante
  Label15.Caption = "Fechas de " & DCTipoComprobante
  Captc = SinEspaciosIzq(DCTipoComprobante.Text)
  Cap1 = Trim$(DCTipoComprobante.Text)
    
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
  If (cod = 4) Or (cod = 5) Then
     FraDctoModificado.Visible = True
     Documento_Modificado
     'Carga en el combo de Documentos Modificados los
     'Tipos de Comprobantes
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo <> 100 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
  Else
     FraDctoModificado.Visible = False
  End If
End Sub

Public Sub Captura_TipoComprobante_DctoModificado()
  CapDcto = Trim$(DCDctoModif.Text)
     
  'Busca que sea igual a la Descripcion
  With AdoTipoComprobante.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & CapDcto & "' ")
       If Not .EOF Then
          CapDm = .Fields("Tipo_Comprobante_Codigo")
       Else
          MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
       End If
    End If
  End With
  If (cod = 4) Or (cod = 5) Then
     FraDctoModificado.Visible = True
     'Verifico si hay documentos modificados de ese Proveedor
     Documento_Modificado
  Else
     FraDctoModificado.Visible = False
  End If
End Sub

Sub Documento_Modificado()
    'Facturas Emitidas del proveedor
     sSQL = "SELECT * " _
          & "FROM Trans_Compras " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND IdProv = '" & CodigoCliente & "' " _
          & "ORDER BY Secuencial "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             CNumSerieTresComp.AddItem .Fields("Secuencial")
            .MoveNext
          Loop
      End If
     End With
End Sub

Public Sub Carga_RetencionIvaBienes_Servicios()
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Bienes <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaBien, AdoRetIvaBienes, sSQL, "Porc"
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Servicios <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaServ, AdoRetIvaServicios, sSQL, "Porc"
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
  SelectDBCombo DCPorcenIva, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDBCombo DCPorcenIce, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  DCConceptoRet = "329 - Por Otros Servicios (N)"
End Sub

Public Sub Limpiar_Controles()
  ac = 0
  DCProveedor.Text = ""
  TxtNumeroC.Text = ""
  DCRetIBienes.Visible = False
  DCRetISer.Visible = False
  ChRetB.value = False
  ChRetF.value = False
  ChRetS.value = False
  LblNumIdent.Caption = ""
  LblTD.Caption = ""
  OpcNo.value = True
  DCSustento.Text = ""
  DCTipoComprobante.Text = ""
  TxtNumSerieUno.Text = ""
  TxtNumSerieDos.Text = ""
  TxtNumSerietres.Text = ""
  TxtNumAutor.Text = ""
  FechaValida MBFechaEmi
  FechaValida MBFechaRegis
  FechaValida MBFechaCad
  TxtBaseImpo.Text = ""
  TxtBaseImpoGrav.Text = ""
  TxtBaseImpoIce.Text = ""
  DCPorcenIva.Text = ""
  TxtMontoIva.Text = ""
  DCPorcenIce.Text = ""
  TxtMontoIce.Text = ""
  TxtIvaBienMonIva.Text = ""
  DCPorcenRetenIvaBien.Text = ""
  TxtIvaBienValRet.Text = ""
  TxtIvaSerMonIva.Text = ""
  DCPorcenRetenIvaServ.Text = ""
  TxtIvaSerValRet.Text = ""
  TxtNumUnoComRet.Text = ""
  TxtNumDosComRet.Text = ""
  TxtNumTresComRet.Text = ""
  TxtNumUnoAutComRet.Text = ""
  TxtSumatoria.Text = ""
  DCConceptoRet.Text = ""
  TxtBimpConA.Text = ""
  TxtPorRetConA.Text = ""
  TxtValConA.Text = ""
  TxtTotalReten.Text = ""
  TxtNumConParPol.Text = ""
  TxtMonTitOner.Text = ""
  TxtMonTitGrat.Text = ""
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
       & "AND Tipo_Trans = 'C' "
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
    Encerar_Var
    Limpiar_Controles
    Listar_Air
   'Cargo el No.Autorización de las retenciones
    TxtNumUnoAutComRet = AutorizaRet
   'Carga el Sustento Tributario
    Carga_CreditoTributario
   'Carga en el Data Combo los Clientes con su RUC
    Leer_Clientes
    DCTipoComprobante.Text = "Factura"
   'Carga la Tabla de Retencion Iva Bienes y Servicios al DataCombo
    Carga_RetencionIvaBienes_Servicios
    DCPorcenIce.Text = ""
   'Carga la Tabla de Conceptos Retencion al DataCombo
    MBFechaRegis = MBFechaEmi
    Carga_ConceptosRetencion MBFechaEmi
    
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RF' " _
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
         & "AND TC = 'RI' " _
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
         & "AND TC = 'RI' " _
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
    LblMensaje.Visible = False
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And Rs And Rb) = 0 Then
           ChRetF.Visible = False
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           LblMensaje.Visible = True
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
       Modificacion_AT "C", CModificacion, CMes, CAño
       CModificacion.Visible = True
       CAño.SetFocus
    End If
End Sub

Public Sub Grabacion()
   'Grabo en el Asiento_Compras e implicito Asiento_Air
    If OpcSi.value = True Then ValorP = "S" Else ValorP = "N"
   'MsgBox CodigoCliente
    FechaTexto = MBFechaRegis
    Factura_No = CTNumero(TxtNumSerietres)
    SetAdoAddNew "Asiento_Compras"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "DevIva", ValorP
    SetAdoFields "CodSustento", Cap
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    SetAdoFields "Secuencial", Factura_No
    SetAdoFields "Autorizacion", TxtNumAutor
    SetAdoFields "FechaEmision", MBFechaEmi
    SetAdoFields "FechaRegistro", MBFechaRegis
    SetAdoFields "FechaCaducidad", MBFechaCad
    SetAdoFields "BaseImponible", CTNumero(TxtBaseImpo, 2)
    SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpoGrav, 2)
    SetAdoFields "PorcentajeIva", CodPorIva
    SetAdoFields "MontoIva", CTNumero(TxtMontoIva, 2)
    SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIce, 2)
    SetAdoFields "PorcentajeIce", CodPorIce
    SetAdoFields "MontoIce", CTNumero(TxtMontoIce, 2)
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBien
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIva, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRet, 2)
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServ
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIva, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRet, 2)
    If (cod = 4) Or (cod = 5) Then
       SetAdoFields "DocModificado", CapDm
       SetAdoFields "FechaEmiModificado", MBFechaEmiComp
       SetAdoFields "EstabModificado", TxtNumSerieUnoComp
       SetAdoFields "PtoEmiModificado", TxtNumSerieDosComp
       SetAdoFields "SecModificado", CNumSerieTresComp
       SetAdoFields "AutModificado", TxtNumAutComp
    Else
       SetAdoFields "DocModificado", "0"
       SetAdoFields "FechaEmiModificado", date
       SetAdoFields "EstabModificado", "000"
       SetAdoFields "PtoEmiModificado", "000"
       SetAdoFields "SecModificado", "0000000"
       SetAdoFields "AutModificado", "0000000000"
    End If
    If TxtNumConParPol = "" Or TxtNumConParPol = "0000000000" Then
       SetAdoFields "ContratoPartidoPolitico", "0000000000"
    Else
       SetAdoFields "ContratoPartidoPolitico", TxtNumConParPol
    End If
    SetAdoFields "MontoTituloOneroso", CTNumero(TxtMonTitOner, 2)
    SetAdoFields "MontoTituloGratuito", CTNumero(TxtMonTitGrat, 2)
   'Verifico si activaron los checks de retenciones
    If ChRetB = 1 Then SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    If ChRetS = 1 Then SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
    SetAdoFields "A_No", Ln_No
    SetAdoFields "T_No", Trans_No
    SetAdoUpdate
    'Grabamos los datos de la transaccion en la tabla definitiva de almacenamiento
    ID_Trans = Maximo_De("Trans_Compras", "ID")  'va a tener el indice de transaccion unico para que no exista duplicados en a base
    sSQL = "SELECT * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "ORDER BY T_No "
    SelectAdodc AdoAsientoCompras, sSQL
    
    With AdoAsientoCompras.Recordset
     If .RecordCount > 0 Then
         'MsgBox "Compras: " & CodigoCliente
         FechaTexto = .Fields("FechaRegistro")
         SetAdoAddNew "Trans_Compras"
         SetAdoFields "T", Normal
         SetAdoFields "IdProv", CodigoCliente
         SetAdoFields "DevIva", .Fields("DevIva")
         SetAdoFields "CodSustento", .Fields("CodSustento")
         SetAdoFields "TipoComprobante", .Fields("TipoComprobante")
         SetAdoFields "Establecimiento", .Fields("Establecimiento")
         SetAdoFields "PuntoEmision", .Fields("PuntoEmision")
         SetAdoFields "Secuencial", .Fields("Secuencial")
         SetAdoFields "Autorizacion", .Fields("Autorizacion")
         SetAdoFields "FechaEmision", .Fields("FechaEmision")
         SetAdoFields "FechaRegistro", .Fields("FechaRegistro")
         SetAdoFields "FechaCaducidad", .Fields("FechaCaducidad")
         SetAdoFields "BaseImponible", .Fields("BaseImponible")
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
         SetAdoFields "TP", CTP
         SetAdoFields "Numero", TxtNumeroC
         SetAdoFields "Fecha", FechaTexto
         SetAdoFields "Porc_Bienes", .Fields("Porc_Bienes")
         SetAdoFields "Porc_Servicios", .Fields("Porc_Servicios")
         SetAdoFields "Cta_Servicio", .Fields("Cta_Servicio")
         SetAdoFields "Cta_Bienes", .Fields("Cta_Bienes")
         SetAdoFields "ID", ID_Trans
         SetAdoFields "Linea_SRI", 0
         SetAdoFields "DocModificado", .Fields("DocModificado")
         SetAdoFields "FechaEmiModificado", .Fields("FechaEmiModificado")
         SetAdoFields "EstabModificado", .Fields("EstabModificado")
         SetAdoFields "PtoEmiModificado", .Fields("PtoEmiModificado")
         SetAdoFields "SecModificado", .Fields("SecModificado")
         SetAdoFields "AutModificado", .Fields("AutModificado")
         SetAdoFields "ContratoPartidoPolitico", .Fields("ContratoPartidoPolitico")
         SetAdoFields "MontoTituloOneroso", .Fields("MontoTituloOneroso")
         SetAdoFields "MontoTituloGratuito", .Fields("MontoTituloGratuito")
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
         & "AND Tipo_Trans = 'C' " _
         & "ORDER BY A_No "
    SelectAdodc AdoTransAir, sSQL
    With AdoAsientoAir.Recordset
     If .RecordCount > 0 Then
         'MsgBox "Air: " & CodigoCliente
         Do While Not .EOF
            SetAdoAddNew "Trans_Air"
            SetAdoFields "T", Normal
            SetAdoFields "IdProv", CodigoCliente
            SetAdoFields "CodRet", .Fields("CodRet")
            SetAdoFields "BaseImp", .Fields("BaseImp")
            SetAdoFields "Porcentaje", .Fields("Porcentaje")
            SetAdoFields "ValRet", .Fields("ValRet")
            SetAdoFields "EstabRetencion", .Fields("EstabRetencion")
            SetAdoFields "PtoEmiRetencion", .Fields("PtoEmiRetencion")
            SetAdoFields "SecRetencion", .Fields("SecRetencion")
            SetAdoFields "AutRetencion", .Fields("AutRetencion")
            SetAdoFields "Tipo_Trans", .Fields("Tipo_Trans")
            SetAdoFields "Numero", TxtNumeroC
            SetAdoFields "TP", CTP
            SetAdoFields "Cta_Retencion", .Fields("Cta_Retencion")
            SetAdoFields "EstabFactura", .Fields("EstabFactura")
            SetAdoFields "PuntoEmiFactura", .Fields("PuntoEmiFactura")
            SetAdoFields "Factura_No", Factura_No
            SetAdoFields "Fecha", FechaTexto
            SetAdoFields "ID", ID_Trans
            SetAdoFields "Linea_SRI", 0
            SetAdoUpdate
            'ID_Trans = ID_Trans + 1
            'Ln_No = Ln_No + 1
           .MoveNext
         Loop
      End If
    End With
End Sub

Public Sub Habilita_Controles()
   'Habilito los controles para la modificacion
    CModificacion.Enabled = True
    SSTCompras.Enabled = True
    DCProveedor.Enabled = True
    CmdGrabar.Enabled = True
    FrmTipoComprob.Enabled = True
    FrmRetencion.Enabled = True
    Label23.Visible = True
    CMes.Visible = True
    CAño.Visible = True
End Sub

Public Sub Deshabilita_Controles()
   'Deshabilito los controles para la modificacion
    CModificacion.Enabled = False
    SSTCompras.Enabled = False
    DCProveedor.Enabled = False
    CmdGrabar.Enabled = False
    FrmTipoComprob.Enabled = False
    FrmRetencion.Enabled = False
    Label23.Visible = False
    CMes.Visible = False
    CAño.Visible = False
End Sub

Public Sub Activar_BS()
    TxtIvaBienMonIva.Enabled = True
    DCPorcenRetenIvaBien.Enabled = True
    TxtIvaBienValRet.Enabled = True
    TxtIvaSerMonIva.Enabled = True
    DCPorcenRetenIvaServ.Enabled = True
    TxtIvaSerValRet.Enabled = True
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Presentamos la malla Asiento Air
  'CodRet,Detalle,BaseImp,Porcentaje,ValRet,EstabRetencion,PtoEmiRetencion,SecRetencion,AutRetencion,FechaEmiRet,Item,CodigoU
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
End Sub

Private Sub TxtPorRetConA_GotFocus()
  MarcarTexto TxtPorRetConA
End Sub

Private Sub TxtPorRetConA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorRetConA_LostFocus()
  If OP = False Then
     TxtValConA = (CTNumero(TxtBimpConA, 2) * CTNumero(TxtPorRetConA, 2)) / 100
     Insertar_AsientoAir
  End If
End Sub

Public Sub Encerar_Var()
  ac = 0
  Ln_No = 0
  DCPorcenIce = 0
  DCPorcenRetenIvaBien = 0
  DCPorcenRetenIvaServ = 0
  CodPorIce = "0"
  CodPorIva = "0"
  CodRetBien = "0"
  CodRetServ = "0"
End Sub
