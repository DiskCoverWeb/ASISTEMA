VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Facturas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FACTURACION:  Ingreso de Facturas"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TBarFactura 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   129
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir de Facturar"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Factura"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar Productos, Marcas y Bodegas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Orden"
            Object.ToolTipText     =   "Asignar Orden de Trabajo"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Guia"
            Object.ToolTipText     =   "Asignar Guia de Remision"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Suscripcion"
            Object.ToolTipText     =   "Asignar Suscripcion/Contrato"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Reserva"
            Object.ToolTipText     =   "Asignar Rserva"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CopiarFactura"
            Object.ToolTipText     =   "Hacer copia de una factura existente"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   750
         Left            =   4830
         TabIndex        =   130
         Top             =   -105
         Width           =   15870
         Begin VB.CheckBox Check1 
            Caption         =   "Facturar en ME"
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
            TabIndex        =   133
            Top             =   210
            Width           =   1695
         End
         Begin VB.CheckBox CheqSP 
            Caption         =   "Sector Público"
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
            TabIndex        =   132
            Top             =   210
            Width           =   1695
         End
         Begin VB.TextBox TxtCompra 
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
            Left            =   5775
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   131
            Text            =   "Facturas.frx":0000
            Top             =   210
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCMod 
            Bindings        =   "Facturas.frx":0004
            DataSource      =   "AdoMod"
            Height          =   360
            Left            =   7245
            TabIndex        =   134
            Top             =   210
            Visible         =   0   'False
            Width           =   6840
            _ExtentX        =   12065
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
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
         Begin VB.Label LabelCodigo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
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
            Height          =   330
            Left            =   14175
            TabIndex        =   136
            Top             =   210
            Width           =   1590
         End
         Begin VB.Label LblCompra 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Orden Compra No."
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
            TabIndex        =   135
            Top             =   210
            Width           =   1800
         End
      End
   End
   Begin VB.Frame FrmCopyFA 
      BackColor       =   &H00FF0000&
      Caption         =   "SERIE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   13440
      TabIndex        =   124
      Top             =   1575
      Visible         =   0   'False
      Width           =   1665
      Begin MSDataListLib.DataCombo DCFACopia 
         Bindings        =   "Facturas.frx":0019
         DataSource      =   "AdoFactura"
         Height          =   2820
         Left            =   105
         TabIndex        =   126
         Top             =   735
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   4974
         _Version        =   393216
         Style           =   1
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
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SELECCIONE LA FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   105
         TabIndex        =   125
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.Frame FrmPVP 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   16905
      TabIndex        =   122
      Top             =   3885
      Visible         =   0   'False
      Width           =   1590
      Begin VB.ListBox LstPVP 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   645
         Left            =   105
         TabIndex        =   123
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.Frame FrmGuiaRemision 
      BackColor       =   &H00FF8080&
      Caption         =   "DATOS DE LA GUIA DE REMISION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   15540
      TabIndex        =   70
      Top             =   4515
      Visible         =   0   'False
      Width           =   8520
      Begin VB.TextBox TxtLugarEntrega 
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   97
         Top             =   3045
         Width           =   4215
      End
      Begin VB.TextBox TxtZona 
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
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   95
         Top             =   3045
         Width           =   1695
      End
      Begin VB.TextBox TxtPedido 
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
         Left            =   1155
         MaxLength       =   15
         TabIndex        =   93
         Top             =   3045
         Width           =   1380
      End
      Begin VB.TextBox TxtPlaca 
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
         Left            =   105
         MaxLength       =   10
         TabIndex        =   91
         Text            =   "XXX-9999"
         Top             =   3045
         Width           =   1065
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF8080&
         Caption         =   "Aceptar"
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
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3465
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FF8080&
         Caption         =   "Cancelar"
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
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3465
         Width           =   1065
      End
      Begin MSMask.MaskEdBox MBoxFechaGRE 
         Height          =   330
         Left            =   3885
         TabIndex        =   72
         Top             =   210
         Width           =   1380
         _ExtentX        =   2434
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
      Begin MSMask.MaskEdBox MBoxFechaGRI 
         Height          =   330
         Left            =   2415
         TabIndex        =   79
         Top             =   1050
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
      Begin MSMask.MaskEdBox MBoxFechaGRF 
         Height          =   330
         Left            =   2415
         TabIndex        =   83
         Top             =   1470
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
      Begin MSDataListLib.DataCombo DCCiudadI 
         Bindings        =   "Facturas.frx":0032
         DataSource      =   "AdoCiudades"
         Height          =   315
         Left            =   4620
         TabIndex        =   81
         Top             =   1050
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "QUITO"
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
      Begin MSDataListLib.DataCombo DCCiudadF 
         Bindings        =   "Facturas.frx":004C
         DataSource      =   "AdoCiudades"
         Height          =   315
         Left            =   4620
         TabIndex        =   85
         Top             =   1470
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "QUITO"
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
      Begin MSDataListLib.DataCombo DCRazonSocial 
         Bindings        =   "Facturas.frx":0066
         DataSource      =   "AdoPersonas"
         Height          =   315
         Left            =   3675
         TabIndex        =   87
         Top             =   1890
         Width           =   4740
         _ExtentX        =   8361
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
      Begin MSDataListLib.DataCombo DCEmpresaEntrega 
         Bindings        =   "Facturas.frx":0080
         DataSource      =   "AdoTransporte"
         Height          =   315
         Left            =   3675
         TabIndex        =   89
         Top             =   2310
         Width           =   4740
         _ExtentX        =   8361
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
      Begin MSDataListLib.DataCombo DCSerieGR 
         Bindings        =   "Facturas.frx":009C
         DataSource      =   "AdoSerieGR"
         Height          =   315
         Left            =   5880
         TabIndex        =   74
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
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
      Begin VB.Label Label51 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Lugar de Entrega"
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
         Left            =   4200
         TabIndex        =   96
         Top             =   2730
         Width           =   4215
      End
      Begin VB.Label LblAutGuiaRem 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3885
         TabIndex        =   77
         Top             =   630
         Width           =   4530
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " AUTORIZACION GUIA DE REMISION"
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
         Left            =   105
         TabIndex        =   76
         Top             =   630
         Width           =   3795
      End
      Begin VB.Label LblGuiaR 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   7035
         TabIndex        =   75
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Zona"
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
         Left            =   2520
         TabIndex        =   94
         Top             =   2730
         Width           =   1695
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Pedido"
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
         Left            =   1155
         TabIndex        =   92
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Placa"
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
         Top             =   2730
         Width           =   1065
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No."
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
         Left            =   5355
         TabIndex        =   73
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Empresa de Transporte"
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
         TabIndex        =   88
         Top             =   2310
         Width           =   3585
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre o Razon Social (Transportista)"
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
         TabIndex        =   86
         Top             =   1890
         Width           =   3585
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
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
         TabIndex        =   84
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
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
         TabIndex        =   80
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Finalizacion del Traslado"
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
         TabIndex        =   82
         Top             =   1470
         Width           =   2325
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Iniciacion del Traslado"
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
         Width           =   2325
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de Emision de la Guia de Remision"
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
         Width           =   3795
      End
   End
   Begin VB.Frame FrmReservas 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DATOS DE LA RESERVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   16905
      TabIndex        =   105
      Top             =   3990
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton Command11 
         Caption         =   " Cancelar"
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
         Left            =   2625
         TabIndex        =   117
         Top             =   1680
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Aceptar"
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
         Left            =   1470
         TabIndex        =   116
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox TxtTipoRooms 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   115
         Top             =   1260
         Width           =   2220
      End
      Begin VB.TextBox TxtCantRooms 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   105
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   113
         Text            =   "Facturas.frx":00B5
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox TxtNoches 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2835
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   111
         Text            =   "Facturas.frx":00B7
         Top             =   525
         Width           =   855
      End
      Begin MSMask.MaskEdBox MBFechaIn 
         Height          =   330
         Left            =   105
         TabIndex        =   107
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin MSMask.MaskEdBox MBFechaOut 
         Height          =   330
         Left            =   1470
         TabIndex        =   109
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VB.Label Label33 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo de Habitación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1470
         TabIndex        =   114
         Top             =   945
         Width           =   2220
      End
      Begin VB.Label Label32 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Noches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2835
         TabIndex        =   112
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label31 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cant. Hab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   110
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label30 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1470
         TabIndex        =   108
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   106
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.TextBox TxtDetalle 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2850
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   41
      Text            =   "Facturas.frx":00B9
      Top             =   4200
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.Frame FrmSeries 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SELECCIONE EL PRODUCTO SEGUN LA SERIE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6090
      TabIndex        =   119
      Top             =   3570
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ListBox LstSeries 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   105
         TabIndex        =   120
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label LblSeries 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRODUCTO"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   4620
      End
   End
   Begin VB.Frame FrmOrdenNo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ORDENES DE PRODUCCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   13545
      TabIndex        =   100
      Top             =   5985
      Visible         =   0   'False
      Width           =   4530
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFF80&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3045
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1680
         Width           =   1380
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFF80&
         Caption         =   "Procesar Seleccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1680
         Width           =   1380
      End
      Begin VB.ListBox LstOrden 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         ItemData        =   "Facturas.frx":00BF
         Left            =   105
         List            =   "Facturas.frx":00C1
         Style           =   1  'Checkbox
         TabIndex        =   102
         Top             =   210
         Width           =   4320
      End
      Begin VB.CommandButton CommandButton1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Imprimir Detalle Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1680
         Width           =   1380
      End
   End
   Begin MSDataListLib.DataCombo DCMedico 
      Bindings        =   "Facturas.frx":00C3
      DataSource      =   "AdoMedico"
      Height          =   360
      Left            =   105
      TabIndex        =   26
      ToolTipText     =   "<Ctrl+R>: Buscar por CI/RUC, <F12>: LLamar a la Historia Clinica, <Ctrl+F>: Listar Ordenes de Trabajo  del Cliente"
      Top             =   2415
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrmFechaV 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha de Venc."
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
      Left            =   13440
      TabIndex        =   46
      Top             =   3990
      Visible         =   0   'False
      Width           =   1590
      Begin MSMask.MaskEdBox MBFechaVGR 
         Height          =   330
         Left            =   105
         TabIndex        =   47
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
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
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "Facturas.frx":00DB
      DataSource      =   "AdoArticulo"
      Height          =   360
      Left            =   3780
      TabIndex        =   40
      ToolTipText     =   "<F10> Insertar Orden de Pedidos"
      Top             =   3570
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   8388608
      Text            =   "Producto"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Facturas.frx":00F5
      DataSource      =   "AdoCliente"
      Height          =   360
      Left            =   1680
      TabIndex        =   17
      ToolTipText     =   "<Ctrl+R>: Buscar por CI/RUC, <F12>: LLamar a la Historia Clinica, <Ctrl+F>: Listar Ordenes de Trabajo  del Cliente"
      Top             =   1575
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSDataListLib.DataCombo DCGrupo_No 
      Bindings        =   "Facturas.frx":010E
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   14280
      TabIndex        =   13
      ToolTipText     =   "<Ctrl-F12> Forzar al 12% IVA"
      Top             =   1155
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Grupo"
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
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   3990
      TabIndex        =   3
      Top             =   735
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
   Begin VB.TextBox TxtEmail 
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
      Left            =   12390
      MaxLength       =   60
      TabIndex        =   25
      Top             =   1995
      Width           =   7785
   End
   Begin VB.ComboBox CDesc1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   14700
      TabIndex        =   49
      Text            =   "Combo1"
      Top             =   3570
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "Facturas.frx":0125
      Height          =   3900
      Left            =   105
      TabIndex        =   69
      Top             =   3990
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   6879
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
         Name            =   "Courier New"
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
      Left            =   16905
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   53
      Text            =   "Facturas.frx":013F
      Top             =   3570
      Width           =   1590
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
      Left            =   15750
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   51
      Text            =   "Facturas.frx":0144
      ToolTipText     =   "<Alt+P> Presenta lista de Precios"
      Top             =   3570
      Width           =   1065
   End
   Begin VB.TextBox TextComEjec 
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
      Left            =   13440
      MaxLength       =   10
      TabIndex        =   45
      Text            =   "0"
      ToolTipText     =   "<Ctrl+F11>:Detalle de Ordenes, <Ctrl+F12>: Detalles de Lotes"
      Top             =   3570
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   16590
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9030
      Width           =   225
   End
   Begin MSDataListLib.DataCombo DCEjecutivo 
      Bindings        =   "Facturas.frx":0146
      DataSource      =   "AdoEjecutivo"
      Height          =   360
      Left            =   10920
      TabIndex        =   28
      Top             =   2415
      Visible         =   0   'False
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3255
      MultiLine       =   -1  'True
      TabIndex        =   61
      Text            =   "Facturas.frx":0161
      Top             =   8925
      Width           =   1485
   End
   Begin VB.TextBox TextComision 
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
      Left            =   19635
      MaxLength       =   5
      TabIndex        =   30
      Text            =   "0"
      Top             =   2415
      Visible         =   0   'False
      Width           =   540
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   1575
      MaxLength       =   200
      TabIndex        =   32
      Top             =   2835
      Width           =   6210
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
      Left            =   8610
      MaxLength       =   200
      TabIndex        =   34
      Top             =   2835
      Width           =   7890
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
      Left            =   18900
      TabIndex        =   9
      Text            =   "000000000"
      Top             =   735
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1365
      TabIndex        =   1
      Top             =   735
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
   Begin VB.CheckBox CheqEjec 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ejecutivo de Venta"
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
      Left            =   8820
      TabIndex        =   27
      Top             =   2415
      Width           =   2010
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2415
      Top             =   5775
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
      Left            =   315
      Top             =   6720
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
      Left            =   315
      Top             =   6090
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
      Left            =   2415
      Top             =   6090
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
      Left            =   315
      Top             =   6405
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
      Left            =   2415
      Top             =   6720
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
      Left            =   315
      Top             =   5775
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
      Height          =   330
      Left            =   2415
      Top             =   6405
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
      Left            =   315
      Top             =   7035
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
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   2415
      Top             =   7035
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
      Left            =   2415
      Top             =   7350
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
   Begin MSAdodcLib.Adodc AdoCorte 
      Height          =   330
      Left            =   315
      Top             =   7350
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
      Caption         =   "Corte"
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
   Begin MSAdodcLib.Adodc AdoOrden 
      Height          =   330
      Left            =   315
      Top             =   7665
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
      Caption         =   "Orden"
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
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "Facturas.frx":0168
      DataSource      =   "AdoBodega"
      Height          =   360
      Left            =   17535
      TabIndex        =   36
      ToolTipText     =   "<F10> Insertar Orden de Pedidos"
      Top             =   2835
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCMarca 
      Bindings        =   "Facturas.frx":0180
      DataSource      =   "AdoMarca"
      Height          =   360
      Left            =   105
      TabIndex        =   38
      ToolTipText     =   "<F10> Insertar Orden de Pedidos"
      Top             =   3570
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   8388608
      Text            =   "."
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   315
      Top             =   7980
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
      Left            =   2415
      Top             =   7665
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
   Begin MSAdodcLib.Adodc AdoCiudades 
      Height          =   330
      Left            =   4515
      Top             =   5775
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
      Caption         =   "Ciudades"
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
   Begin MSAdodcLib.Adodc AdoPersonas 
      Height          =   330
      Left            =   4515
      Top             =   6090
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
      Caption         =   "Personas"
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
   Begin MSDataListLib.DataCombo DCTipoPago 
      Bindings        =   "Facturas.frx":0197
      DataSource      =   "AdoTipoPago"
      Height          =   360
      Left            =   1680
      TabIndex        =   11
      Top             =   1155
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   2415
      Top             =   7980
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
      Caption         =   "TipoPago"
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
   Begin MSAdodcLib.Adodc AdoSerieGR 
      Height          =   330
      Left            =   4515
      Top             =   6405
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
      Caption         =   "SerieGR"
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
      Bindings        =   "Facturas.frx":01B1
      DataSource      =   "AdoLinea"
      Height          =   360
      Left            =   6930
      TabIndex        =   5
      Top             =   735
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "CxC Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPorcIVA 
      Height          =   330
      Left            =   4515
      Top             =   7035
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
      Caption         =   "PorcIVA"
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
   Begin MSDataListLib.DataCombo DCPorcIVA 
      Bindings        =   "Facturas.frx":01C8
      DataSource      =   "AdoPorcIVA"
      Height          =   360
      Left            =   13335
      TabIndex        =   7
      ToolTipText     =   "Seleccione el Porc del IVA"
      Top             =   735
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "12%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoTransporte 
      Height          =   330
      Left            =   4515
      Top             =   7350
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
      Caption         =   "Transporte"
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
   Begin MSAdodcLib.Adodc AdoMedico 
      Height          =   330
      Left            =   4515
      Top             =   6720
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
      Caption         =   "Medico"
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
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Utilidad Neta"
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
      TabIndex        =   128
      Top             =   8610
      Width           =   1590
   End
   Begin VB.Label LblUtilidad 
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
      Left            =   9660
      TabIndex        =   127
      Top             =   8925
      Width           =   1590
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Porc IVA"
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
      Left            =   12390
      TabIndex        =   6
      Top             =   735
      Width           =   960
   End
   Begin VB.Label LblGuia 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   645
      Left            =   11550
      TabIndex        =   118
      Top             =   8610
      Width           =   5265
   End
   Begin VB.Label Label27 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta x Cobrar"
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
      TabIndex        =   4
      Top             =   735
      Width           =   1590
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   15855
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":01E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":04FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":0815
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":0B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":0E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":1163
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":147D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Facturas.frx":1797
            Key             =   ""
         EndProperty
      EndProperty
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
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   18585
      TabIndex        =   55
      Top             =   3570
      Width           =   1590
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO DE PAGO"
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
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label LabelStockArt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   39
      Top             =   3255
      Width           =   8415
   End
   Begin VB.Label LabelTelefono 
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
      Height          =   330
      Left            =   16905
      TabIndex        =   21
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label LabelRUC 
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
      Height          =   330
      Left            =   13860
      TabIndex        =   19
      Top             =   1575
      Width           =   2010
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telefono:"
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
      Left            =   15960
      TabIndex        =   20
      Top             =   1575
      Width           =   960
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I./R.U.C."
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
      Left            =   12390
      TabIndex        =   18
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
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
      Top             =   1575
      Width           =   1590
   End
   Begin VB.Label Label21 
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
      Height          =   330
      Left            =   9450
      TabIndex        =   23
      Top             =   1995
      Width           =   1380
   End
   Begin VB.Label Label24 
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
      Height          =   330
      Left            =   105
      TabIndex        =   22
      Top             =   1995
      Width           =   9360
   End
   Begin VB.Label LblSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999,999.99"
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
      Left            =   18585
      TabIndex        =   15
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente del Cliente"
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
      Left            =   15960
      TabIndex        =   14
      Top             =   1155
      Width           =   2640
   End
   Begin VB.Label Label48 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CORREO"
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
      TabIndex        =   24
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Label Label38 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ESCOJA &GRUPO"
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
      Left            =   12390
      TabIndex        =   12
      Top             =   1155
      Width           =   1905
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MARCA"
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
      TabIndex        =   37
      Top             =   3255
      Width           =   3585
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA"
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
      Left            =   16590
      TabIndex        =   35
      Top             =   2835
      Width           =   960
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
      Left            =   18585
      TabIndex        =   54
      Top             =   3255
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.V.P."
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
      Left            =   16905
      TabIndex        =   52
      Top             =   3255
      Width           =   1590
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
      Left            =   15750
      TabIndex        =   50
      Top             =   3255
      Width           =   1065
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desc. %"
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
      Left            =   14700
      TabIndex        =   48
      Top             =   3255
      Width           =   960
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ord./Lote"
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
      Left            =   13440
      TabIndex        =   44
      Top             =   3255
      Width           =   1170
   End
   Begin VB.Label LabelStock 
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
      Left            =   12285
      TabIndex        =   43
      Top             =   3570
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stock"
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
      Left            =   12285
      TabIndex        =   42
      Top             =   3255
      Width           =   1065
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vencimiento"
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
      TabIndex        =   2
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label LabelTotal 
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
      Left            =   7980
      TabIndex        =   67
      Top             =   8925
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
      Left            =   7980
      TabIndex        =   66
      Top             =   8610
      Width           =   1590
   End
   Begin VB.Label LabelIVA 
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
      TabIndex        =   65
      Top             =   8925
      Width           =   1485
   End
   Begin VB.Label LabelServ 
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
      Left            =   4830
      TabIndex        =   63
      Top             =   8925
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
      Left            =   6405
      TabIndex        =   64
      Top             =   8610
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
      Left            =   4830
      TabIndex        =   62
      Top             =   8610
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
      Left            =   3255
      TabIndex        =   60
      Top             =   8610
      Width           =   1485
   End
   Begin VB.Label LabelConIVA 
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
      Left            =   1680
      TabIndex        =   59
      Top             =   8925
      Width           =   1485
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision %"
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
      Left            =   18375
      TabIndex        =   29
      Top             =   2415
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " OBSERVACION"
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
      TabIndex        =   31
      Top             =   2835
      Width           =   1485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOTA"
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
      Left            =   7875
      TabIndex        =   33
      Top             =   2835
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 0000000000000 NOTA DE VENTA No. 001001-"
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
      Left            =   14280
      TabIndex        =   8
      Top             =   735
      Width           =   4635
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Emision"
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
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label LabelSubTotal 
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
      Left            =   105
      TabIndex        =   57
      Top             =   8925
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
      Left            =   1680
      TabIndex        =   58
      Top             =   8610
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
      TabIndex        =   56
      Top             =   8610
      Width           =   1485
   End
End
Attribute VB_Name = "Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'mlucas/1210
'7300002003
Dim XProducto As String
Dim AnchoDetalle As Single
Dim CodArtOrden As String
Dim Ln_No_O As Byte
Dim Valor_UnitA As Currency
Dim Terminar_FA As Boolean
Dim Lote_No As String
Dim No_Hab As String
Dim StockLote As Currency

Public Sub Tipo_De_Facturacion()
  If FA.TC = "NV" Then
     Facturas.Caption = "INGRESAR NOTA DE VENTA"
     Label2.Caption = " " & FA.Autorizacion & " NOTA DE VENTA No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. 0.00%"
  ElseIf FA.TC = "OP" Then
     Facturas.Caption = "INGRESAR ORDEN DE PEDIDO"
     Label2.Caption = " " & FA.Autorizacion & " ORDEN No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. 0.00%"
  Else
     Facturas.Caption = "INGRESAR FACTURA"
     Label2.Caption = " " & FA.Autorizacion & " FACTURA No. " & FA.Serie & "-"
     Label3.Caption = "I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
  End If
  'Facturas.Caption = Facturas.Caption & " (" & FA.TC & ")"
  Label36.Caption = "Serv. " & Format$(Porc_Serv * 100, "#0.00") & "%"
  TipoFactura = FA.TC
End Sub

Public Sub Grabar_Factura_Actual()
Dim GuiaRemision As Long
   'Procedemos a grabar la factura actual
    FechaValida MBoxFechaV
    TextoValido TextObs
    TextoValido TextNota
    TextoValido TxtPedido
    TextoValido TxtCompra, True, , 0
    TextoValido TxtZona, , True
    TextoValido TextComision, , True
    TextoValido TxtLugarEntrega, , True

    If FA.CodigoC = Ninguno Then
       MsgBox "Error: No se puede grabar factura sin nombre de Cliente o Beneficiario"
    Else
       sSQL = "SELECT " & Full_Fields("Asiento_F") & " " _
            & "FROM Asiento_F " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' " _
            & "ORDER BY A_No "
       SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|TOTAL 4|."
       Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
       If AdoAsientoF.Recordset.RecordCount > 0 Then
          RatonReloj
          
          Calculos_Totales_Factura FA
          
          LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
          LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
          TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
          LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
          LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
          LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
    
          RatonNormal
          Titulo = "FORMULARIO DE GRABACION"
          Mensajes = "Esta Seguro que desea grabar: " & vbCrLf
          If FA.TC = "OP" Then
             Mensajes = Mensajes & "La Orden de Producción No. " & TextFacturaNo
             TipoFactura = "OP"
          Else
             Mensajes = Mensajes & "La Factura No. " & TextFacturaNo
          End If
          
          If BoxMensaje = vbYes Then
             Moneda_US = False
             FA.Nuevo_Doc = True
             FA.Factura = Val(TextFacturaNo)
             TextoFormaPago = PagoCred
             FA.Tipo_Pago = SinEspaciosIzq(DCTipoPago)
             If Val(FA.Tipo_Pago) <= 0 Then FA.Tipo_Pago = "01"
             If Check1.value = 1 Then Moneda_US = True
             If CheqSP.value = 1 Then FA.SP = True
             FA.T = Pendiente
             FA.Fecha_V = MBoxFechaV
             FA.Orden_Compra = "0"
             FA.SubCta = Ninguno
             FA.Porc_IVA = Porc_IVA
             FA.Forma_Pago = TextoFormaPago
             FA.Observacion = TextObs
             FA.Nota = TextNota
             FA.Pedido = TxtPedido
            'MsgBox Val(TxtCompra)
             If IsNumeric(TxtCompra) Then FA.Orden_Compra = Format(Val(TxtCompra), "0000000000")
             If AdoMod.Recordset.RecordCount > 0 Then
                AdoMod.Recordset.MoveFirst
                AdoMod.Recordset.Find ("Detalle = '" & DCMod.Text & "' ")
                If Not AdoMod.Recordset.EOF Then FA.SubCta = AdoMod.Recordset.fields("Codigo")
             End If
             FA.ME_ = Moneda_US
             FA.Saldo_MN = FA.Total_MN
             Comision = Redondear(Val(TextComision) / 100, 4)
             Total_Comision = Redondear(Total_SubTotal * Comision, 2)
            
            'Datos del Encabezado y totales de la factura
             FA.Porc_C = Comision
             FA.Comision = Total_Comision
            '--------------------------------------------------------------------------------------------------------------------
             If Existe_Factura(FA) Then
                Titulo = "FORMULARIO DE CONFIRMACION"
                Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
                         & "Ya existe " & FA.TC & " No. " & FA.Serie & "-" & Format$(FA.Factura, "000000000") & vbCrLf & vbCrLf _
                         & "Desea Reprocesarla"
                If BoxMensaje = vbYes Then FA.Nuevo_Doc = False Else GoTo NoGrabarFA
             Else
                Factura_No = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
                If FA.Factura <> Factura_No Then
                   Titulo = "Formulario de Confirmación"
                   Mensajes = "La " & FA.TC & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") & ", no esta Procesada, desea Procesarla?"
                   If BoxMensaje = vbYes Then FA.Nuevo_Doc = False Else GoTo NoGrabarFA
                End If
             End If
            'MsgBox FA.CodigoC
            '------------------
            'Grabamos el numero de factura
             Grabar_Factura FA, True
             
             SRI_Autorizacion.Tipo_Doc_SRI = TipoDoc
            '-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-
            ' If Ambiente = "2" Then SRI_Enviar_Mails FA, SRI_Autorizacion
            '-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-.-.-.--.-
            'Grabamos Abonos del numero de factura
             RatonNormal
             Bandera = False
             Evaluar = True
             FechaTexto = MBoxFecha
             Factura_No = FA.Factura
             Numero = Factura_No
             Titulo = "Formulario de Grabacion"
             If FA.TC = "OP" Then
                Mensajes = "Anticipo de Abonos"
                TipoFactura = "OP"
                If BoxMensaje = vbYes Then AbonoAnticipado.Show 1
             Else
                Mensajes = "Pago al Contado"
                If BoxMensaje = vbYes Then Abonos.Show 1
             End If
            'MsgBox "Desktop Test: Documento " & FA.TC & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") & vbCrLf & "Guia No. " & FA.Serie_GR & "-" & FA.Remision
             RatonReloj
            'Autorizamos la factura y/o Guia de Remision
             If Len(FA.Autorizacion) = 13 Then SRI_Crear_Clave_Acceso_Facturas FA, False, , True
               'MsgBox "Desktop Test: Documento " & FA.TC & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000")
                If Len(FA.Autorizacion_GR) = 13 Then
                   SRI_Crear_Clave_Acceso_Guia_Remision FA, False, True
                  'MsgBox "Desktop Test: Guia No. " & FA.TC & " No. " & FA.Serie_GR & "-" & Format(FA.Remision, "000000000")
                   If Len(FA.Autorizacion_GR) > 13 Then
                      sSQL = "UPDATE Facturas_Auxiliares " _
                           & "SET Fecha_Aut_GR = #" & BuscarFecha(FA.Fecha_Aut_GR) & "#," _
                           & "Autorizacion_GR = '" & FA.Autorizacion_GR & "'," _
                           & "Clave_Acceso_GR = '" & FA.ClaveAcceso_GR & "'," _
                           & "Estado_SRI_GR = '" & FA.Estado_SRI_GR & "'," _
                           & "Hora_Aut_GR = '" & FA.Hora_GR & "' " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND TC = '" & FA.TC & "' " _
                           & "AND Serie = '" & FA.Serie & "' " _
                           & "AND Factura = " & FA.Factura & " " _
                           & "AND Autorizacion = '" & FA.Autorizacion & "' "
                      Ejecutar_SQL_SP sSQL
                   End If
                End If
                TA.TP = FA.TC
                TA.Serie = FA.Serie
                TA.Factura = FA.Factura
                TA.Autorizacion = FA.Autorizacion
                TA.CodigoC = FA.CodigoC
                
                Actualizar_Saldos_Facturas_SP TA.TP, TA.Serie, TA.Factura
                
                RatonNormal
                'MsgBox "..."
                If FA.TC <> "OP" Then
                  'MsgBox FA.Autorizacion & vbCrLf & FA.Autorizacion_GR
                   If Len(FA.Autorizacion) >= 13 Then
                      Imprimir_Punto_Venta FA
                   Else
                      Titulo = "IMPRESION"
                      Mensajes = "Facturacion Multiple"
                      If BoxMensaje = vbYes Then
                         Factura_Desde = FA.Factura
                         Factura_Hasta = FA.Factura
                         FA.Tipo_PRN = "FM"
                         Imprimir_Facturas_CxC Facturas, FA, True
                      Else
                         FA.Tipo_PRN = "FA"
                         Imprimir_Facturas FA
                      End If
                   End If
                   Facturas_Impresas FA
                End If
                RatonReloj
                If FA.TC <> "OP" Then
                   If FA.Remision > 0 Then
                      If Len(FA.Autorizacion_GR) < 13 Then
                         Imprimir_Guia_Remision AdoFactura, AdoAsientoF, FA
                      ElseIf Len(FA.Autorizacion_GR) >= 13 Then
                         SRI_Generar_PDF_GR FA, True
                      End If
                   End If
                End If
                SRI_Generar_PDF_FA FA, True
                LblGuiaR.Caption = "0"
                LblGuia.Caption = ""
                CheqSP.value = 0
                Ln_No = 0
                Total_Desc = 0
                Encerar_Factura FA
                FA.Nuevo_Doc = True
                FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
                TextFacturaNo.Text = FA.Factura
                sSQL = "SELECT * " _
                     & "FROM Asiento_F " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND CodigoU = '" & CodigoUsuario & "' "
                Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
                RatonNormal
                DCLinea.SetFocus
             Else
                RatonNormal
                MsgBox "Revise los datos ingresados y vuelva a intentar grabar"
             End If
       Else
NoGrabarFA:
           RatonNormal
           MsgBox "No se procedio a grabar el documento " & FA.TC & " No. " & FA.Serie & "-" _
           & Format(FA.Factura, "000000000") & ", revise los datos ingresados y vuelva a intentar"
       End If
  End If
End Sub

Public Sub DatosArticulos()
Dim EsNumero As Boolean
  With AdoArticulo.Recordset
       Producto = DatInv.Producto
       Cta_Ventas = DatInv.Cta_Ventas
       
       TextVUnit.Text = Format$(DatInv.PVP, "#,##0.0000")
       NumStrg = TextVUnit.Text
       If DatInv.IVA Then NumStrg = Format$(DatInv.PVP + (DatInv.PVP * Porc_IVA), "#,##0." & String(Dec_PVP, "0"))
       LabelStockArt.Caption = " P R O D U C T O              " & Space(20 - Len(NumStrg)) & Moneda & " " & NumStrg
       VUnitAnterior = DatInv.PVP
       LabelStock.Caption = DatInv.Stock
       Codigos = DatInv.Codigo_Inv
       CodigoInv1 = DatInv.Codigo_Barra
       BanIVA = DatInv.IVA
       If TipoFactura = "NV" Then BanIVA = False
       DCArticulo.Text = Producto
       TextComEjec.Text = "0"
       'TxtDetalle.SetFocus
          TxtDetalle.Text = Producto
          If Len(DatInv.Detalle) > 3 Then TxtDetalle.Text = TxtDetalle.Text & vbCrLf & DatInv.Detalle
          EsNumero = False
          If IsNumeric(DatInv.Codigo_Barra) Then
             If Val(DatInv.Codigo_Barra) > 0 Then EsNumero = True
          End If
          If Len(DatInv.Codigo_Barra) > 1 And EsNumero Then TxtDetalle.Text = TxtDetalle.Text & vbCrLf & "S/N: " & DatInv.Codigo_Barra
          TxtDetalle.Visible = True
          'TxtDetalle.SetFocus
  End With
End Sub

Private Sub CDesc1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CDesc1_LostFocus()
  No_Hab = Ninguno
  If Val(CDesc1.Text) > 0 Then
     No_Hab = InputBox("INGRESE DETALLE ADICIONAL(20)", "DETALLE ADICIONAL", "")
     No_Hab = UCaseStrg(MidStrg(No_Hab, 1, 40))
  End If
End Sub

Private Sub CheqEjec_Click()
 FA.Cod_Ejec = Ninguno
 If CheqEjec.value = 1 Then
    DCEjecutivo.Visible = True
    Label11.Visible = True
    TextComision.Visible = True
 Else
    DCEjecutivo.Visible = False
    Label11.Visible = False
    TextComision.Visible = False
 End If
End Sub

Private Sub CheqEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CheqSP_Click()
   If CheqSP.value = 1 Then
      FA.SP = True
      TxtCompra.SetFocus
   Else
      FA.SP = False
   End If
End Sub

Private Sub Command1_Click()
  Unload Facturas
End Sub

Private Sub Command10_Click()
  TxtDetalle.Text = Producto
  If Len(DatInv.Detalle) > 3 Then TxtDetalle.Text = TxtDetalle.Text & vbCrLf & DatInv.Detalle
  TextCant = TxtNoches
  FrmReservas.Visible = False
  TxtDetalle.Visible = True
  TxtDetalle.SetFocus
End Sub

Private Sub Command11_Click()
  FrmReservas.Visible = False
  DCArticulo.SetFocus
End Sub

Private Sub Command4_Click()
   Llenar_Orden LstOrden
   FrmOrdenNo.Visible = False
   TextComEjec = Lote_No
   TextComEjec.SetFocus
End Sub

Private Sub Command5_Click()
    DatInv.Reg_Sanitario = Ninguno
    DatInv.Procedencia = Ninguno
    DatInv.Modelo = Ninguno
    'DatInv.Serie_No = Ninguno
    DatInv.Fecha_Exp = FechaSistema
    DatInv.Fecha_Fab = FechaSistema
    TextComEjec.SetFocus
    FrmOrdenNo.Visible = False
End Sub

Private Sub Command7_Click()
  FA.FechaGRE = FechaSistema
  FA.FechaGRI = FechaSistema
  FA.FechaGRF = FechaSistema
  FA.CiudadGRI = Ninguno
  FA.CiudadGRF = Ninguno
  FA.Comercial = Ninguno
  FA.CIRUCComercial = Ninguno
  FA.Zona = Ninguno
  FA.Entrega = Ninguno
  FA.Lugar_Entrega = Ninguno
  FA.CIRUCEntrega = Ninguno
  FA.Dir_EntregaGR = Ninguno
  FA.Placa_Vehiculo = Ninguno
  FA.Autorizacion_GR = Ninguno
  FA.ClaveAcceso_GR = Ninguno
  FA.Serie_GR = Ninguno
  FA.Remision = 0
  LblGuia.Caption = ""
' Command13.SetFocus
  FrmGuiaRemision.Visible = False
End Sub

Private Sub Command8_Click()
    FA.ClaveAcceso_GR = Ninguno
    FA.Autorizacion_GR = LblAutGuiaRem.Caption
    FA.Serie_GR = DCSerieGR
    FA.Remision = Val(LblGuiaR.Caption)
    FA.FechaGRE = MBoxFechaGRE
    FA.FechaGRI = MBoxFechaGRI
    FA.FechaGRF = MBoxFechaGRF
    FA.Placa_Vehiculo = TxtPlaca
    FA.Lugar_Entrega = TxtLugarEntrega
    FA.Zona = TxtZona
    If FA.Remision > 0 Then
       With AdoCiudades.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Descripcion_Rubro = '" & DCCiudadI & "' ")
            If Not .EOF Then FA.CiudadGRI = .fields("Descripcion_Rubro")
           .MoveFirst
           .Find ("Descripcion_Rubro = '" & DCCiudadF & "' ")
            If Not .EOF Then FA.CiudadGRF = .fields("Descripcion_Rubro")
        End If
       End With
       With AdoPersonas.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Cliente = '" & DCRazonSocial & "' ")
            If Not .EOF Then
               FA.Comercial = .fields("Cliente")
               FA.CIRUCComercial = .fields("CI_RUC")
               FA.Dir_PartidaGR = .fields("Direccion")
            End If
        End If
       End With
       
       With AdoTransporte.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Cliente = '" & DCEmpresaEntrega & "' ")
            If Not .EOF Then
               FA.Entrega = .fields("Cliente")
               FA.CIRUCEntrega = .fields("CI_RUC")
               FA.Dir_EntregaGR = .fields("Direccion")
            End If
        End If
       End With
    End If
   'Command13.SetFocus
    LblGuia.Caption = "Guia de Remision: " & FA.Serie_GR & "-" & Format(FA.Remision, "000000000") & vbCrLf _
                    & "Autorizacion: " & FA.Autorizacion_GR
    FrmGuiaRemision.Visible = False
End Sub

Private Sub CommandButton1_Click()
   Orden_No = Val(InputBox("Imprimir el Detalle" & vbCrLf & "de la Orden No.", "IMPRESION DE ORDEN DE TRABAJO", "0"))
   sSQL = "SELECT Fecha,Producto,Cantidad,Precio,A,L,S " _
        & "FROM Trans_Ticket " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Ticket = " & Orden_No & " " _
        & "AND TC = 'OP' " _
        & "ORDER BY Producto "
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Mensajes = "Imprimir Orden de Trabajo"
        Titulo = "IMPRESION"
        MensajeEncabData = "LISTA DE ORDEN DE TRABAJO No. " & Format$(Orden_No, "000000")
        SQLMsg1 = "Cliente: " & DCCliente
        Cuadricula = True
        If BoxMensaje = vbYes Then ImprimirAdo AdoAux, True, 1, 8, True
    End If
   End With
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCEjecutivo_LostFocus()
   With AdoEjecutivo.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente Like '" & DCEjecutivo & "' ")
        If Not .EOF Then
           FA.Cod_Ejec = .fields("Codigo")
           TextComision = Format$(.fields("Porc_Com") * 100, "#0.00")
        Else
           MsgBox "Ejecutivo de Venta no asignado"
           FA.Cod_Ejec = Ninguno
        End If
    Else
        MsgBox "Ejecutivo de Venta no asignado"
        FA.Cod_Ejec = Ninguno
    End If
   End With
End Sub

Private Sub DCEmpresaEntrega_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCEmpresaEntrega
    If Len(Busqueda) > 0 Then
       sSQL = "SELECT TOP 50 Cliente, CI_RUC, Codigo, Cta_CxP, Grupo, Cod_Ejec, Direccion " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       sSQL = sSQL _
            & "AND TD IN ('C','R') " _
            & "ORDER BY Cliente "
       Select_Adodc AdoTransporte, sSQL
    End If
End Sub

Private Sub DCGrupo_No_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PorcIva As Byte
 Keys_Especiales Shift
 PresionoEnter KeyCode
 If CtrlDown And KeyCode = vbKeyF12 Then
    PorcIva = InputBox("Ingrese el porcentaje a Proccesar:", "PORCENTAJE DE IVA", Porc_IVA * 100)
    Select Case PorcIva
      Case 8, 10, 12, 14, 15
           Porc_IVA = PorcIva / 100
      Case Else
           Porc_IVA = 0.12
    End Select
    Tipo_De_Facturacion
 End If
End Sub

Private Sub DCGrupo_No_LostFocus()
    Label10.Caption = " CLIENTES "
    If DCGrupo_No = "" Then DCGrupo_No = Ninguno
    Grupo_No = DCGrupo_No
    Listar_Tipo_Beneficiarios Grupo_No
   'MsgBox AdoCliente.Recordset.RecordCount
    Label10.Caption = " CLIENTES (" & AdoCliente.Recordset.RecordCount & "):"
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  Encerar_Factura FA
  FA.Cod_CxC = DCLinea
  FA.Serie = Ninguno
  FA.Autorizacion = Ninguno
  Lineas_De_CxC FA
   Tipo_De_Facturacion
   FechaTexto1 = MBoxFecha
   FA.Fecha = MBoxFecha
   FA.Fecha_V = FA.Fecha
   FA.Fecha_C = FA.Fecha
   FechaComp = FA.Fecha
   
  'FA.Factura = Numero_Factura(FA)
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  TextFacturaNo = FA.Factura
  DatInv.TC = FA.TC
End Sub

Private Sub DCArticulo_GotFocus()
  Terminar_FA = False
 'DCArticulo.width = TextCant.Left - DCArticulo.Left
  LabelStock.Caption = ""
  DatInv.Reg_Sanitario = Ninguno
  DatInv.Procedencia = Ninguno
  DatInv.Modelo = Ninguno
  DatInv.Fecha_Exp = FechaSistema
  DatInv.Fecha_Fab = FechaSistema
  DatInv.Serie_No = Ninguno
  Lote_No = Ninguno
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Patron As String
  XProducto = DCArticulo
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then
     Empleados = False
     Calculos_Totales_Factura FA
     Terminar_FA = True
     Grabar_Factura_Actual
  End If
  If KeyCode = vbKeyF10 Then
         Mensajes = "LLENAR LOS DATOS DE" & vbCrLf _
                  & "LA FACTURA/NOTA DE VENTA" & vbCrLf _
                  & "DESDE CERO"
         Titulo = "FORMULARIO DE ELIMINACION"
         If BoxMensaje = vbYes Then
            Ln_No = 0
            sSQL = "DELETE * " _
                 & "FROM Asiento_F " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND CodigoU = '" & CodigoUsuario & "' "
            Ejecutar_SQL_SP sSQL
         End If
         Cadena = DCCliente.Text & vbCrLf & vbCrLf _
                & "ORDEN No."
         Habitacion_No = UCaseStrg(InputBox(Cadena, "FACTURACION DE PEDIDOS", ""))
         If Habitacion_No = "" Then Habitacion_No = Ninguno
''''         sSQL = "SELECT Codigo,Producto,Cta_Venta,SUM(Cantidad) As Cant,AVG(PRECIO) As PVP,SUM(Total) As VTotal,SUM(Total_IVA) As VTotal_IVA " _
''''              & "FROM Trans_Pedidos " _
''''              & "WHERE Item = '" & NumEmpresa & "' " _
''''              & "AND No_Hab = '" & Habitacion_No & "' " _
''''              & "GROUP BY Codigo,Producto,Cta_Venta " _
''''              & "ORDER BY Codigo "
''''         Select_Adodc AdoAux, sSQL
''''         With AdoAux.Recordset
''''          If .RecordCount > 0 Then
''''              Do While Not .EOF
''''                 Codigo = .Fields("Codigo")
''''                 Codigo1 = .Fields("Producto")
''''                 Cta = .Fields("Cta_Venta")
''''                 Precio = .Fields("PVP")
''''                 Total = .Fields("VTotal")
''''                 Total_IVA = .Fields("VTotal_IVA")
''''                 Cantidad = .Fields("Cant")
''''                 Insertar_Pedidos
''''                .MoveNext
''''              Loop
''''          End If
''''         End With
         sSQL = "SELECT TP.Codigo_Sup,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0," _
              & "SUM(TP.Cantidad) As Cant,AVG(TP.PRECIO) As PVP,SUM(TP.Total) As VTotal," _
              & "SUM(TP.Total_IVA) As VTotal_IVA " _
              & "FROM Trans_Pedidos As TP,Catalogo_Productos As CP " _
              & "WHERE TP.Item = '" & NumEmpresa & "' " _
              & "AND TP.Periodo = '" & Periodo_Contable & "' " _
              & "AND TP.No_Hab = '" & Habitacion_No & "' " _
              & "AND TP.Codigo_Sup = CP.Codigo_Inv " _
              & "AND TP.Item = CP.Item " _
              & "AND TP.Periodo = CP.Periodo " _
              & "GROUP BY TP.Codigo_Sup,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0 " _
              & "ORDER BY TP.Codigo_Sup "
         Select_Adodc AdoAux, sSQL
         With AdoAux.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Codigo = .fields("Codigo_Sup")
                 Codigo1 = .fields("Producto")
                 If FA.TC = "NV" Then
                    Cta = .fields("Cta_Ventas_0")
                 Else
                    Cta = .fields("Cta_Ventas")
                 End If
                 Precio = .fields("PVP")
                 Total = .fields("VTotal")
                 Total_IVA = .fields("VTotal_IVA")
                 Cantidad = .fields("Cant")
                 Insertar_Pedidos
                .MoveNext
              Loop
          End If
         End With
         sSQL = "SELECT TP.Codigo,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0," _
              & "SUM(TP.Cantidad) As Cant,AVG(TP.PRECIO) As PVP,SUM(TP.Total) As VTotal," _
              & "SUM(TP.Total_IVA) As VTotal_IVA " _
              & "FROM Trans_Pedidos As TP,Catalogo_Productos As CP " _
              & "WHERE TP.Item = '" & NumEmpresa & "' " _
              & "AND TP.Periodo = '" & Periodo_Contable & "' " _
              & "AND TP.No_Hab = '" & Habitacion_No & "' " _
              & "AND CP.Agrupacion = " & Val(adFalse) & " " _
              & "AND TP.Codigo = CP.Codigo_Inv " _
              & "AND TP.Item = CP.Item " _
              & "AND TP.Periodo = CP.Periodo " _
              & "GROUP BY TP.Codigo,CP.Producto,CP.Cta_Ventas,CP.Cta_Ventas_0 " _
              & "ORDER BY TP.Codigo "
         Select_Adodc AdoAux, sSQL
         With AdoAux.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Codigo = .fields("Codigo")
                 Codigo1 = .fields("Producto")
                 If FA.TC = "NV" Then
                    Cta = .fields("Cta_Ventas_0")
                 Else
                    Cta = .fields("Cta_Ventas")
                 End If
                 Precio = .fields("PVP")
                 Total = .fields("VTotal")
                 Total_IVA = .fields("VTotal_IVA")
                 Cantidad = .fields("Cant")
                 Insertar_Pedidos
                .MoveNext
              Loop
          End If
         End With
         sSQL = "SELECT No_Hab,Fecha,Hora,Producto,Cantidad,Precio,Total,Total_IVA " _
              & "FROM Trans_Pedidos " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND No_Hab = '" & Habitacion_No & "' " _
              & "ORDER BY Codigo "
         Select_Adodc AdoAux, sSQL
         With AdoAux.Recordset
          If .RecordCount > 0 Then
              Mensajes = "Imprimir Pedidos"
              Titulo = "IMPRESION"
              MensajeEncabData = "LISTA DE PEDIDOS A FACTURAR"
              SQLMsg1 = "Cliente: " & DCCliente
              If BoxMensaje = vbYes Then ImprimirAdo AdoAux, True, 1, 9
          End If
         End With
         sSQL = "SELECT * " _
              & "FROM Asiento_F " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND CodigoU = '" & CodigoUsuario & "' "
         SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|TOTAL 4|."
         Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
         Calculos_Totales_Factura FA
         LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
         LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
         TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
         LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
         LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
         LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
         DCArticulo = XProducto
         DCBodega.SetFocus
         TxtDetalle.Visible = False
  End If
  If CtrlDown And KeyCode = vbKeyF Then
      sSQL = "DELETE * " _
           & "FROM Asiento_F " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' "
      Ejecutar_SQL_SP sSQL
      If CodigoCliente <> "" Then
         sSQL = "SELECT Cliente,CP.Producto,CP.IVA,TF.* " _
              & "FROM Clientes As C,Catalogo_Productos As CP,Trans_Fletes As TF " _
              & "WHERE TF.Item = '" & NumEmpresa & "' " _
              & "AND TF.Periodo = '" & Periodo_Contable & "' " _
              & "AND TF.CodigoC = '" & CodigoCliente & "' " _
              & "AND TF.Ok <> " & Val(adFalse) & " " _
              & "AND TF.T = '" & Normal & "' " _
              & "AND C.Codigo = TF.CodigoC " _
              & "AND CP.Item = TF.Item " _
              & "AND CP.Periodo = TF.Periodo " _
              & "AND CP.Codigo_Inv = TF.Codigo_Inv " _
              & "ORDER BY TF.Fecha_I "
         Select_Adodc AdoAux, sSQL
         With AdoAux.Recordset
          If .RecordCount > 0 Then
              Do While Not .EOF
                 Real3 = 0
                 If .fields("IVA") Then Real3 = Redondear(.fields("Flete") * Porc_IVA, 4)
                 CodigoP = Format$(.fields("Numero"), "0000000")
                 Producto = .fields("Fecha_I") & Space(2) & .fields("Ruta") _
                          & Space(20 - Len(.fields("Ruta"))) & .fields("Referencia") _
                          & Space(10 - Len(.fields("Referencia"))) & .fields("Carga") _
                          & Space(19 - Len(.fields("Carga"))) & CodigoP
                 SetAdoAddNew "Asiento_F"
                 SetAdoFields "CODIGO_L", CodigoL
                 SetAdoFields "CODIGO", .fields("Codigo_Inv")
                 SetAdoFields "PRODUCTO", Producto
                 SetAdoFields "CANT", 1
                 SetAdoFields "PRECIO", .fields("Flete")
                 SetAdoFields "TOTAL", .fields("Flete")
                 SetAdoFields "Total_IVA", Real3
                 SetAdoFields "Cta", Cta_Ventas
                 SetAdoFields "Numero", .fields("Numero")
                 SetAdoFields "Item", NumEmpresa
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoUpdate
                .MoveNext
              Loop
          End If
         End With
      End If
      sSQL = "SELECT * " _
           & "FROM Asiento_F " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' "
      Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
      Calculos_Totales_Factura FA
   End If
   If CtrlDown And KeyCode = vbKeyB Then
      Patron = InputBox("INGRESE EL PATRON DE BUSQUEDA:", "PATRON DE BUSQUEDA")
      Listar_Productos DCArticulo, AdoArticulo, OpcServicio, Patron
      DCArticulo.SetFocus
   End If
   If CtrlDown And KeyCode = vbKeyS Then
      LstSeries.Clear
      sSQL = "SELECT CP.Codigo_Inv,CP.Producto,TK.Serie_No,SUM(Entrada-Salida) As Stock " _
           & "FROM Catalogo_Productos As CP, Trans_Kardex As TK " _
           & "WHERE CP.Item = '" & NumEmpresa & "' " _
           & "AND CP.Periodo = '" & Periodo_Contable & "' " _
           & "AND CP.Producto = '" & XProducto & "' " _
           & "AND CP.Item = TK.Item " _
           & "AND CP.Periodo = TK.Periodo " _
           & "AND CP.Codigo_Inv = TK.Codigo_Inv " _
           & "GROUP BY CP.Codigo_Inv,CP.Producto,TK.Serie_No " _
           & "HAVING SUM(Entrada-Salida) > 0 " _
           & "ORDER BY CP.Producto,TK.Serie_No "
      Select_Adodc AdoAux, sSQL
      With AdoAux.Recordset
       If .RecordCount > 0 Then
           Codigo1 = " S E R I E  No."
           Codigo1 = Codigo1 & String(30 - Len(Codigo1), " ")
           Codigo2 = "EXISTENCIA"
           LblSeries.Caption = Codigo1 & " " & Codigo2
           Do While Not .EOF
              Codigo1 = .fields("Serie_No")
              Codigo1 = Codigo1 & String(30 - Len(Codigo1), " ")
              Codigo2 = Format(.fields("Stock"), "#,##0.00")
              Codigo2 = String(10 - Len(Codigo2), " ") & Codigo2
              LstSeries.AddItem Codigo1 & " " & Codigo2
             .MoveNext
           Loop
           CentrarFrame FrmSeries
           FrmSeries.Visible = True
           LstSeries.Text = LstSeries.List(0)
           LstSeries.SetFocus
       Else
           MsgBox "Este producto no tiene existencia"
           DCArticulo.SetFocus
       End If
      End With
   End If
End Sub

Private Sub DCArticulo_LostFocus()
  If Not Terminar_FA Then
     Codigos = Ninguno
     Cod_Marca = Ninguno
     Cod_Bodega = Ninguno
     With AdoBodega.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Bodega Like '" & DCBodega & "' ")
          If Not .EOF Then Cod_Bodega = .fields("CodBod")
      End If
     End With
     Cod_Marca = Ninguno
     With AdoMarca.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Marca Like '" & DCMarca & "' ")
          If Not .EOF Then Cod_Marca = .fields("CodMar")
      End If
     End With
     If Leer_Codigo_Inv(DCArticulo, FechaSistema, Cod_Bodega, Cod_Marca) Then DatosArticulos
     LstPVP.Clear
     LstPVP.AddItem Format(DatInv.PVP, "#,##0." & String(Dec_PVP, "0"))
     If DatInv.PVP2 > 0 Then LstPVP.AddItem Format(DatInv.PVP2, "#,##0." & String(Dec_PVP, "0"))
     If DatInv.PVP3 > 0 Then LstPVP.AddItem Format(DatInv.PVP3, "#,##0." & String(Dec_PVP, "0"))
     LstPVP.Text = LstPVP.List(0)
  End If
'''   If DatInv.Stock <= 0 And Len(DatInv.Cta_Inventario) > 1 Then
'''      MsgBox "Usted no puede Facturar este producto, no tiene Stock"
'''      DCArticulo.SetFocus
'''   End If
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyF2 Then
     Codigos = InputBox("Historia Clinica:", "CODIGO DE HISTORIA CLINICA", "0")
     With AdoCliente.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Actividad Like '" & Codigos & "' ")
          If Not .EOF Then
             CodigoC = .fields("Cliente")
             SiguienteControl
          End If
      End If
     End With
  End If
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 0 Then
       sSQL = "SELECT TOP 50 Cliente, CI_RUC, Codigo, Cta_CxP, Grupo, Cod_Ejec " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '" & Busqueda & "%' "
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoCliente, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
  CodigoCliente = Ninguno
  NombreCliente = Ninguno
  DireccionCli = Ninguno
  LabelCodigo.Caption = Ninguno
  LabelTelefono.Caption = Ninguno
  LabelRUC.Caption = Ninguno
  LblSaldo.Caption = "0.00"
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       If IsNumeric(DCCliente.Text) Then
         .Find ("CI_RUC = '" & DCCliente.Text & "' ")
       Else
         .Find ("Cliente = '" & DCCliente.Text & "' ")
       End If
       If Not .EOF Then
          CodigoCliente = .fields("Codigo")
          TBeneficiario = Leer_Datos_Cliente_SP(CodigoCliente)
          FA.CodigoC = TBeneficiario.Codigo
          FA.Cliente = TBeneficiario.Cliente
          FA.TD = TBeneficiario.TD
          FA.CI_RUC = TBeneficiario.CI_RUC
          FA.TelefonoC = TBeneficiario.Telefono1
          FA.DireccionC = TBeneficiario.Direccion
          FA.EmailC = TBeneficiario.Email1
          FA.EmailC2 = TBeneficiario.Email2
          FA.EmailR = TBeneficiario.EmailR
          
          CodigoCliente = FA.CodigoC
          NombreCliente = FA.Cliente
          LabelTelefono.Caption = FA.TelefonoC
          LabelCodigo.Caption = CodigoCliente
          LabelRUC.Caption = FA.CI_RUC
          DireccionCli = TBeneficiario.Direccion
          DireccionGuia = TBeneficiario.Direccion
          Label21.Caption = " No. " & TBeneficiario.Actividad
          Label24.Caption = " Dir: " & TBeneficiario.Direccion
          NoDias = TBeneficiario.Credito
          TxtEmail = FA.EmailC
          LblSaldo.Caption = Format$(TBeneficiario.Saldo_Pendiente, "#,##0.00")
          If NoDias > 0 Then MBoxFechaV.Text = CLongFecha(CFechaLong(MBoxFecha.Text) + NoDias)
          Label13.Caption = " C.I./R.U.C. (" & FA.TD & ")"
          
          'SiguienteControl
          'If Mod_Fact Then TextFacturaNo.SetFocus Else TextObs.SetFocus
          If Len(.fields("Cta_CxP")) > 1 Then
            'DCGrupo_No = TBeneficiario.Grupo_No
             DCEjecutivo.Text = Ninguno
             FA.Cta_CxP = .fields("Cta_CxP")
             FA.Cod_Ejec = .fields("Cod_Ejec")
             If ComisionEjec Then
                If AdoEjecutivo.Recordset.RecordCount > 0 Then
                   AdoEjecutivo.Recordset.MoveFirst
                   AdoEjecutivo.Recordset.Find ("Codigo = '" & FA.Cod_Ejec & "' ")
                   If Not AdoEjecutivo.Recordset.EOF Then DCEjecutivo.Text = AdoEjecutivo.Recordset.fields("Cliente")
                End If
             End If
             If AdoLinea.Recordset.RecordCount > 0 Then
                AdoLinea.Recordset.MoveFirst
                AdoLinea.Recordset.Find ("CxC = '" & FA.Cta_CxP & "' ")
                If Not AdoLinea.Recordset.EOF Then DCLinea.Text = AdoLinea.Recordset.fields("Concepto")
             End If
             FA.Cod_CxC = DCLinea
          End If
       Else
          Nuevo = True
          NombreCliente = DCCliente.Text
          Facturas.Visible = False
          MsgBox "Cliente no Asignado"
          NivelNo = DCGrupo_No
          FClientesFlash.Show 1
          Facturas.Visible = True
          DCGrupo_No.SetFocus
       End If
      '-------------------
      'MsgBox FA.CodigoC
      '-------------------
   Else
       'MsgBox "No existen datos"
       Nuevo = True
       NombreCliente = DCCliente.Text
       Facturas.Visible = False
       MsgBox "Cliente no Asignado"
       FClientesFlash.Show 1
       Facturas.Visible = True
       DCGrupo_No.SetFocus
   End If
  End With
End Sub

Private Sub DCMarca_LostFocus()
  Cod_Marca = Ninguno
  With AdoMarca.Recordset
   If .RecordCount > 0 Then
      .Find ("Marca = '" & DCMarca & "' ")
       If Not .EOF Then Cod_Marca = .fields("CodMar")
   End If
  End With
  If Len(DCMarca) > 1 Then
     Listar_Productos DCArticulo, AdoArticulo, OpcServicio, , DCMarca
  Else
     Listar_Productos DCArticulo, AdoArticulo, OpcServicio
  End If
End Sub

Private Sub DCMedico_GotFocus()
  FA.CodigoDr = Ninguno
End Sub

Private Sub DCMedico_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCMedico_LostFocus()
   With AdoMedico.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente = '" & DCMedico & "' ")
        If Not .EOF Then
           FA.CodigoDr = .fields("Codigo")
        Else
           MsgBox "Nombre incorrecto"
        End If
    End If
   End With
End Sub

Private Sub DCMod_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCMod_LostFocus()
  FA.SubCta = Ninguno
  If AdoMod.Recordset.RecordCount > 0 Then
     AdoMod.Recordset.MoveFirst
     AdoMod.Recordset.Find ("Detalle = '" & DCMod.Text & "' ")
     If Not AdoMod.Recordset.EOF Then FA.SubCta = AdoMod.Recordset.fields("Codigo")
  End If
  'DCMod.Visible = False
  MBoxFecha.SetFocus
End Sub

Private Sub DCPorcIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCPorcIVA_LostFocus()
  If IsNumeric(DCPorcIVA.Text) Then
     If AdoPorcIVA.Recordset.RecordCount > 0 Then Porc_IVA = Redondear(DCPorcIVA / 100, 2)
  Else
     Porc_IVA = 0
  End If
  Tipo_De_Facturacion
End Sub

Private Sub DCRazonSocial_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCRazonSocial
    If Len(Busqueda) > 0 Then
       sSQL = "SELECT TOP 50 Cliente, CI_RUC, Codigo, Cta_CxP, Grupo, Cod_Ejec, Direccion " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       sSQL = sSQL _
            & "AND TD IN ('C','R') " _
            & "ORDER BY Cliente "
       Select_Adodc AdoPersonas, sSQL
    End If
End Sub

Private Sub DCSerieGR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerieGR_LostFocus()
   LblGuiaR = ReadSetDataNum("GR_SERIE_" & DCSerieGR, True, False)
   
'''   sSQL = "SELECT Serie_GR,MAX(Remision) As GRemision " _
'''        & "FROM Facturas_Auxiliares " _
'''        & "WHERE Item = '" & NumEmpresa & "' " _
'''        & "AND Periodo = '" & Periodo_Contable & "' " _
'''        & "AND Serie_GR = '" & DCSerieGR & "' " _
'''        & "AND Remision > 0 " _
'''        & "GROUP BY Serie_GR"
'''   Select_Adodc AdoAux, sSQL
'''   If AdoAux.Recordset.RecordCount > 0 Then
'''      LblGuiaR = Format(AdoAux.Recordset.Fields("GRemision") + 1, "000000000")
'''   Else
'''      LblGuiaR = "000000001"
'''   End If
   If AdoSerieGR.Recordset.RecordCount > 0 Then
      AdoSerieGR.Recordset.MoveFirst
      AdoSerieGR.Recordset.Find ("Serie = '" & DCSerieGR & "' ")
      If Not AdoSerieGR.Recordset.EOF Then
         LblAutGuiaRem.Caption = AdoSerieGR.Recordset.fields("Autorizacion")
      Else
         LblAutGuiaRem.Caption = ""
      End If
   End If
End Sub

Private Sub DGAsientoF_AfterDelete()
 'Calculos_Totales_Factura Facturas, AdoAsientoF
End Sub

Private Sub DGAsientoF_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & Chr(13) & "(" _
           & AdoAsientoF.Recordset.fields("CODIGO") & ") " _
           & AdoAsientoF.Recordset.fields("PRODUCTO") & "   TOTAL -> " _
           & AdoAsientoF.Recordset.fields("TOTAL") & "?"
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
       If Not .EOF And .fields("Producto") <> Ninguno Then
          TxtDetalle.Visible = True
'          TxtDetalle.Text = DGAsientoF.Columns(1) & ": " & vbCrLf & .Fields("Detalle")
          TxtDetalle.SetFocus
       End If
   End If
  End With
End Sub

Private Sub DGAsientoF_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ID_Trans_Aux As Long
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then TxtDetalle.Visible = False
  If CtrlDown And KeyCode = vbKeyR Then
     With AdoAsientoF.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          Do While Not .EOF
            'SubTotal
             SubTotal = Redondear(.fields("CANT") * .fields("PRECIO"), 2)
            'Descuento
             SubTotalDescuento = .fields("Total_Desc")
            'IVA = SubTotal - Descuento
             If .fields("Total_IVA") > 0 Then SubTotalIVA = Redondear((SubTotal - SubTotalDescuento) * Porc_IVA, 2)
            .fields("Total_IVA") = SubTotalIVA
            .fields("TOTAL") = Redondear(SubTotal, 2)
            .fields("VALOR_TOTAL") = Redondear(SubTotal + SubTotalIVA, 2)
            .Update
            .MoveNext
          Loop
      End If
     End With
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|Total_IVA 4|."
     Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
     Calculos_Totales_Factura FA
     LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
     LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
     TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
     LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
     LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
     LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
     RatonNormal
     MsgBox "Proceso recalculado exitosamente"
  End If
  If AltDown And KeyCode = vbKeyM Then
     'MsgBox "....."
  End If
End Sub

Private Sub DCFACopia_KeyDown(KeyCode As Integer, Shift As Integer)
Dim MesN As Integer
Dim Anio As String
Dim MesS As String

  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then
     FrmCopyFA.Visible = False
     DCLinea.SetFocus
  End If
  If KeyCode = vbKeyReturn Then
     FrmCopyFA.Visible = False
     If AdoFactura.Recordset.RecordCount > 0 Then
        RatonReloj
        Factura_No = DCFACopia.Text
        Ln_No = 1
        sSQL = "DELETE * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Ejecutar_SQL_SP sSQL
        MesN = Month(MBoxFecha.Text)
        Anio = Year(MBoxFecha.Text)
        MesS = MesesLetras(MesN)

        sSQL = "INSERT INTO Asiento_F (Codigo_Cliente, CODIGO, CODIGO_L, PRODUCTO, CANT, PRECIO, TOTAL, Total_Desc, Total_Desc2, Total_IVA, Serie_No, CodBod, COSTO, Cta, Item, CodigoU, " _
             & "A_No, Mes, NoMes, TICKET, CANT_BONIF, Fecha_IN, Fecha_OUT, Cant_Hab, Tipo_Hab, Orden_No, Cod_Ejec, Porc_C, HABIT, RUTA, CodMar, TONELAJE, CORTE, PRECIO2, COD_BAR, " _
             & "Fecha_V, Lote_No, Fecha_Fab, Fecha_Exp, Modelo, Procedencia, Cmds, Utilidad) " _
             & "SELECT CodigoC, Codigo, '" & FA.Cod_CxC & "', Producto, Cantidad, Precio, Total, Total_Desc, Total_Desc2, Total_IVA, Serie_No, CodBodega, Costo, '" & FA.Cta_CxP & "', " _
             & "'" & NumEmpresa & "', '" & CodigoUsuario & "', ROW_NUMBER() OVER(ORDER BY ID ASC), '" & MesS & "', " & MesN & ", '" & Anio & "', Cant_Bonif, Fecha_IN, Fecha_OUT, Cant_Hab, " _
             & "Tipo_Hab, Orden_No, Cod_Ejec, Porc_C, No_Hab, Ruta, CodMarca, Tonelaje, Corte, Precio2, Codigo_Barra, Fecha_V, Lote_No, Fecha_Fab, Fecha_Exp, Modelo, Procedencia, Cmds, " _
             & "CASE WHEN Costo > 0 THEN ROUND(Total-(Cantidad*Costo),2,0) ELSE 0 END " _
             & "FROM Detalle_Factura " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TC = '" & FA.TC & "' " _
             & "AND Serie = '" & FA.Serie & "' " _
             & "AND Factura = " & Factura_No & " " _
             & "ORDER BY ID,Codigo "
'   Clipboard.Clear
'   Clipboard.SetText sSQL
             
        Ejecutar_SQL_SP sSQL
        Eliminar_Nulos_SP "Asiento_F"
     End If
     DGAsientoF.Visible = True
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|TOTAL 4|."
     Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
     Calculos_Totales_Factura FA
     If AdoAsientoF.Recordset.RecordCount > 0 Then
        AdoAsientoF.Recordset.MoveFirst
        Ln_No = AdoAsientoF.Recordset.RecordCount + 1
        sSQL = "SELECT Cliente, CI_RUC, Codigo, Cta_CxP, Grupo, Cod_Ejec " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & AdoAsientoF.Recordset.fields("Codigo_Cliente") & "' "
        Select_Adodc AdoCliente, sSQL
        If AdoCliente.Recordset.RecordCount > 0 Then DCCliente.Text = AdoCliente.Recordset.fields("Cliente") Else DCCliente.Text = Ninguno
     End If
     LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
     LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
     TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
     LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
     LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
     LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
     
     DCCliente.SetFocus
  End If
End Sub

Private Sub Form_Activate()
   FechaValida MBoxFecha, True
   FechaValida MBoxFechaV, False
   FechaValida MBFechaVGR, False
   Mifecha = BuscarFecha(FechaSistema)
   Grupo_No = Ninguno
   Lote_No = Ninguno
   Fecha_Exp = Ninguno
   Fecha_Fab = Ninguno
   Reg_Sanitario = Ninguno
   NombreMarca = Ninguno
   StockLote = 0
   LblGuiaR.Caption = "0"
   LblSaldo.Caption = "0.00"
   LabelCodigo.Caption = ""
   Label10.Caption = " CLIENTES"
   DCMod.Visible = False
   DCMedico.Visible = False
  
   Facturas.WindowState = 2
   DGAsientoF.width = MDI_X_Max - 100
   DGAsientoF.Height = MDI_Y_Max - DGAsientoF.Top - 900
   DGAsientoF.Refresh
   
   Label22.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LabelSubTotal.Top = Label22.Top + Label22.Height + 10
   
   Label23.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LabelConIVA.Top = Label22.Top + Label22.Height + 10
   
   Label6.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   TextDesc.Top = Label22.Top + Label22.Height + 10
   
   Label36.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LabelServ.Top = Label22.Top + Label22.Height + 10
   
   Label3.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LabelIVA.Top = Label22.Top + Label22.Height + 10
   
   Label26.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LabelTotal.Top = Label22.Top + Label22.Height + 10
   
   Label52.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LblUtilidad.Top = Label22.Top + Label22.Height + 10
   
   LblGuia.Top = DGAsientoF.Top + DGAsientoF.Height + 60
   LblGuia.Caption = ""
   
  sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
       & "AND Codigo IN ('01','16','17','18','19','20','21') " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
  
  sSQL = "SELECT CC.Codigo " _
       & "FROM Catalogo_Cuentas As CC INNER JOIN Catalogo_Productos As CP " _
       & "ON CC.Item = CP.Item " _
       & "AND CC.Periodo = CP.Periodo " _
       & "AND CC.Codigo IN (CP.Cta_Ventas,CP.Cta_Ventas_0,CP.Cta_Ventas_Ant) " _
       & "AND CC.Item = '" & NumEmpresa & "' " _
       & "AND CC.Periodo = '" & Periodo_Contable & "' " _
       & "AND CC.DG = 'D' " _
       & "AND CC.TC IN ('I','CC') "
  Select_Adodc AdoAux, sSQL

  sSQL = "SELECT Detalle, Codigo, TC " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If AdoAux.Recordset.RecordCount > 0 Then
     sSQL = sSQL & "AND TC IN ('I','CC') "
  Else
     sSQL = sSQL & "AND TC = 'NN' "
  End If
  sSQL = sSQL & "ORDER BY Detalle "
  SelectDB_Combo DCMod, AdoMod, sSQL, "Detalle"
  If AdoMod.Recordset.RecordCount > 0 Then DCMod.Visible = True
  
  sSQL = "SELECT TOP 50 Cliente, CI_RUC, TD, Codigo, Cta_CxP " _
       & "FROM Clientes " _
       & "WHERE Asignar_Dr <> " & Val(adFalse) & " " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCMedico, AdoMedico, sSQL, "Cliente"
  If AdoMedico.Recordset.RecordCount > 0 Then DCMedico.Visible = True
   
  sSQL = "SELECT Grupo " _
       & "FROM Clientes " _
       & "WHERE T = 'N' " _
       & "AND FA <> " & Val(adFalse) & " " _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDB_Combo DCGrupo_No, AdoGrupo, sSQL, "Grupo"
  
  FA.TC = TipoFactura
  FA.Fecha = MBoxFecha
  sSQL = "SELECT Codigo, Concepto, CxC, Serie, Autorizacion " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact = '" & FA.TC & "' " _
       & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
       & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
  
 'MsgBox FA.Fecha & vbCrLf & DCLinea.Text
  FA.Cod_CxC = DCLinea
  CantFact = 1
  PorCodigo = CBool(ReadSetDataNum("PorCodigo", True, False))
  Lineas_De_CxC FA
  Tipo_De_Facturacion
  TextCant.Text = "0"
  TextVUnit.Text = "0"
  LabelVTotal.Caption = "0"
  Modificar = False
  Bandera = True
  
 'MsgBox TipoFactura
  CDesc1.Clear
  CDesc1.AddItem "00.00"
  sSQL = "SELECT * " _
       & "FROM Catalogo_Interes " _
       & "WHERE TP = 'D' " _
       & "ORDER BY Interes "
  Select_Adodc AdoCorte, sSQL
  With AdoCorte.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           CDesc1.AddItem Format$(.fields("Interes"), "00.00")
          .MoveNext
        Loop
    End If
   End With
   CDesc1.Text = "00.00"
''   sSQL = "SELECT * " _
''        & "FROM Tabla_Costos " _
''        & "WHERE Item = '" & NumEmpresa & "' " _
''        & "ORDER BY Concepto "
''   Select_Adodc AdoCorte, sSQL
    
   TipoDoc = TipoFactura
             
   If ComisionEjec Then
      sSQL = "SELECT CR.Codigo,C.Cliente,C.CI_RUC,CR.Porc_Com " _
           & "FROM Catalogo_Rol_Pagos As CR, Clientes As C " _
           & "WHERE CR.Item = '" & NumEmpresa & "' " _
           & "AND CR.Periodo = '" & Periodo_Contable & "' " _
           & "AND CR.Codigo = C.Codigo " _
           & "ORDER BY C.Cliente "
      SelectDB_Combo DCEjecutivo, AdoEjecutivo, sSQL, "Cliente"
   Else
      DCEjecutivo.Text = Ninguno
   End If
   sSQL = "SELECT CodBod, Bodega " _
        & "FROM Catalogo_Bodegas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Bodega "
   SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
   If AdoBodega.Recordset.RecordCount <= 0 Then
      Label28.Visible = False
      DCBodega.Visible = False
   End If
   sSQL = "SELECT CodMar, Marca " _
        & "FROM Catalogo_Marcas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Marca "
   SelectDB_Combo DCMarca, AdoMarca, sSQL, "Marca"
   If AdoMarca.Recordset.RecordCount <= 0 Then
      LabelStockArt.width = LabelStockArt.width + Label29.width + Label29.Left
      DCArticulo.width = DCArticulo.width + DCMarca.width + DCMarca.Left
      LabelStockArt.Left = Label29.Left
      DCArticulo.Left = DCMarca.Left
      Label29.Visible = False
      DCMarca.Visible = False
   End If
   
   sSQL = "SELECT Codigo, Porc " _
        & "FROM Tabla_Por_ICE_IVA " _
        & "WHERE IVA <> " & Val(adFalse) & " " _
        & "AND Fecha_Inicio <= #" & BuscarFecha(FechaSistema) & "# " _
        & "AND Fecha_Final >= #" & BuscarFecha(FechaSistema) & "# " _
        & "ORDER BY Porc DESC "
   SelectDB_Combo DCPorcIVA, AdoPorcIVA, sSQL, "Porc"
    
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "DELETE * " _
        & "FROM Asiento_TK " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   
  ' Listar_Tipo_Beneficiarios DCGrupo_No
   
   sSQL = "SELECT * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|Total_IVA " & CStr(Dec_IVA) & "|TOTAL 2|."
   Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
   
   TextFacturaNo.Enabled = Mod_Fact
   CheqEjec.Visible = ComisionEjec
'   If ComisionEjec Then CheqEjec.Visible = True Else CheqEjec.Visible = False
   If NombreUsuario = "Administrador de Red" Then
      Command4.Enabled = True
      TextFacturaNo.Enabled = True
   End If
   Total_Desc = 0
   Ln_No = 0
   
   Listar_Productos DCArticulo, AdoArticulo, OpcServicio
   
   Listar_Lotes
   
   sSQL = "SELECT " & Full_Fields("Catalogo_Lineas") & " " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fact = 'OP' "
   Select_Adodc AdoAux, sSQL
'   If AdoAux.Recordset.RecordCount > 0 Then Command9.Visible = True
  'If FA.TC = "OP" Then Command9.Visible = False
   'AnchoDetalle = DCArticulo.width
   If Bloquear_Control Then Command2.Enabled = False
   If ComisionEjec Then
      CheqEjec.value = 1
      DCEjecutivo.Visible = True
      Label11.Visible = True
      TextComision.Visible = True
   End If
   CheqSP.value = 0
   CheqSP.Visible = 0
   If Leer_Campo_Empresa("SP") Then
      CheqSP.Visible = 1
      If LstOrden.ListCount > 0 Then CheqSP.value = 1 Else CheqSP.value = 0
   End If
   RatonNormal
   Check1.SetFocus
End Sub

Private Sub Form_Deactivate()
   Facturas.WindowState = 1
End Sub

Private Sub Form_Load()
  'CentrarForm Facturas
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoMod
   ConectarAdodc AdoCorte
   ConectarAdodc AdoOrden
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoLinea
   ConectarAdodc AdoMarca
   ConectarAdodc AdoBodega
   ConectarAdodc AdoPorcIVA
   ConectarAdodc AdoFactura
   ConectarAdodc AdoAsientoF
   ConectarAdodc AdoCliente
   ConectarAdodc AdoListFact
   ConectarAdodc AdoArticulo
   ConectarAdodc AdoTipoPago
   ConectarAdodc AdoEjecutivo
   ConectarAdodc AdoCiudades
   ConectarAdodc AdoPersonas
   ConectarAdodc AdoTransporte
   ConectarAdodc AdoMedico
   ConectarAdodc AdoSerieGR
   
   SRI_Obtener_Datos_Comprobantes_Electronicos
End Sub

Private Sub LstOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyA Then
     If ClaveSupervisor Then
        IR = Val(InputBox("Activar la Orden No.", "ACTIVAR ORDER DE PRODUCCION", "0"))
        If IR <> 0 Then
           sSQL = "UPDATE Facturas " _
                & "SET T = 'P' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & IR & " " _
                & "AND TC = 'OP' "
           Ejecutar_SQL_SP sSQL
           sSQL = "UPDATE Detalle_Factura " _
                & "SET T = 'P' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & IR & " " _
                & "AND TC = 'OP' "
           Ejecutar_SQL_SP sSQL
           sSQL = "UPDATE Trans_Abonos " _
                & "SET T = 'P' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & IR & " " _
                & "AND TP = 'OP' "
           Ejecutar_SQL_SP sSQL
           Listar_Ordenes
           DLOrden.Text = IR
        End If
     End If
  End If
  
  If KeyCode = vbKeyEscape Then
     FrmOrdenNo.Visible = False
    'Command9.SetFocus
  End If
End Sub

Private Sub LstPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then FrmPVP.Visible = False
End Sub

Private Sub LstPVP_LostFocus()
  TextVUnit.Text = LstPVP.Text
  FrmPVP.Visible = False
  TextVUnit.SetFocus
End Sub

Private Sub LstSeries_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then
     FrmSeries.Visible = False
     'DatInv.Serie_No = Ninguno
     DCArticulo.SetFocus
  End If
  If KeyCode = vbKeyReturn Then
     FrmSeries.Visible = False
     DatInv.Serie_No = SinEspaciosIzq(LstSeries.Text)
    'MsgBox LstSeries.Text & vbCrLf & vbCrLf & DatInv.Serie_No
     TxtDetalle.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyS Then
     DatInv.Serie_No = InputBox("Ingrese la Serie manualmente", "INGRESO DE SERIE MANUAL", ".")
     DatInv.Serie_No = Trim(DatInv.Serie_No)
     If DatInv.Serie_No = "" Then DatInv.Serie_No = Ninguno
     FrmSeries.Visible = False
     TxtDetalle.SetFocus
  End If
End Sub

Private Sub MBFechaIn_GotFocus()
  MarcarTexto MBFechaIn
End Sub

Private Sub MBFechaIn_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaIn_LostFocus()
  FechaValida MBFechaIn
End Sub

Private Sub MBFechaOut_GotFocus()
  MarcarTexto MBFechaOut
End Sub

Private Sub MBFechaOut_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaOut_LostFocus()
  FechaValida MBFechaOut
  TxtNoches = Format((CFechaLong(MBFechaOut) - CFechaLong(MBFechaIn)) + 1, "#0.00")
End Sub

Private Sub MBFechaVGR_GotFocus()
  MarcarTexto MBFechaVGR
End Sub

Private Sub MBFechaVGR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaVGR_LostFocus()
  FechaValida MBFechaVGR, False
  FrmFechaV.Visible = False
  CDesc1.SetFocus
End Sub

Private Sub MBoxFecha_GotFocus()
  MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
    FechaValida MBoxFecha, True
    FA.Fecha = MBoxFecha
    If FA.TC = "OP" Then
       MBoxFechaV = CLongFecha(CFechaLong(MBoxFecha) + 3)
    Else
       MBoxFechaV = CLongFecha(CFechaLong(MBoxFecha) + 15)
    End If
    
    sSQL = "SELECT Codigo, Concepto, CxC, Serie, Autorizacion " _
         & "FROM Catalogo_Lineas " _
         & "WHERE TL <> " & Val(adFalse) & " " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Fact = '" & FA.TC & "' " _
         & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
         & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
         & "ORDER BY Codigo "
    SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
    
    sSQL = "SELECT Codigo, Porc " _
         & "FROM Tabla_Por_ICE_IVA " _
         & "WHERE IVA <> " & Val(adFalse) & " " _
         & "AND Fecha_Inicio <= #" & BuscarFecha(FA.Fecha) & "# " _
         & "AND Fecha_Final >= #" & BuscarFecha(FA.Fecha) & "# " _
         & "ORDER BY Porc DESC "
    SelectDB_Combo DCPorcIVA, AdoPorcIVA, sSQL, "Porc"
End Sub

Private Sub MBoxFechaGRE_GotFocus()
  MarcarTexto MBoxFechaGRE
End Sub

Private Sub MBoxFechaGRE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaGRE_LostFocus()
   FechaValida MBoxFechaGRE, True
   
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fact = 'GR' " _
        & "AND #" & BuscarFecha(MBoxFechaGRE) & "# BETWEEN Fecha and Vencimiento " _
        & "ORDER BY Serie "
   SelectDB_Combo DCSerieGR, AdoSerieGR, sSQL, "Serie"
End Sub

Private Sub MBoxFechaGRI_GotFocus()
  MarcarTexto MBoxFechaGRI
End Sub

Private Sub MBoxFechaGRI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaGRI_LostFocus()
   FechaValida MBoxFechaGRI, True
End Sub

Private Sub MBoxFechaGRF_GotFocus()
  MarcarTexto MBoxFechaGRF
End Sub

Private Sub MBoxFechaGRF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaGRF_LostFocus()
   FechaValida MBoxFechaGRF, True
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV
  FA.Fecha_V = MBoxFechaV
  If (CFechaLong(MBoxFechaV) - CFechaLong(MBoxFecha)) > 0 Then
     TextObs = "CREDITO A " & (CFechaLong(MBoxFechaV) - CFechaLong(MBoxFecha)) & " DIA(S)."
  Else
     TextObs = "CONTADO."
  End If
End Sub

Private Sub TBarFactura_ButtonClick(ByVal Button As ComctlLib.Button)
  'MsgBox Button.key
   Select Case Button.key
    Case "Salir"
          RatonNormal
          Unload Facturas
     Case "Grabar"
          Grabar_Factura_Actual
     Case "Actualizar"
          RatonReloj
          sSQL = "SELECT * " _
               & "FROM Catalogo_Bodegas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY Bodega "
          SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
          
          sSQL = "SELECT * " _
               & "FROM Catalogo_Marcas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY Marca "
          SelectDB_Combo DCMarca, AdoMarca, sSQL, "Marca"
          Listar_Productos DCArticulo, AdoArticulo, OpcServicio
          RatonNormal
     Case "Orden"
          Listar_Ordenes
     Case "Guia"
          sSQL = "SELECT Descripcion_Rubro " _
               & "FROM Tabla_Naciones " _
               & "WHERE TR = 'C' " _
               & "ORDER BY Descripcion_Rubro "
          SelectDB_Combo DCCiudadI, AdoCiudades, sSQL, "Descripcion_Rubro"
          SelectDB_Combo DCCiudadF, AdoCiudades, sSQL, "Descripcion_Rubro"
            
          sSQL = "SELECT TOP 50 Cliente,CI_RUC,TD,Direccion,Codigo " _
               & "FROM Clientes " _
               & "WHERE TD IN ('C','R') " _
               & "ORDER BY Cliente "
          SelectDB_Combo DCRazonSocial, AdoPersonas, sSQL, "Cliente"
          SelectDB_Combo DCEmpresaEntrega, AdoTransporte, sSQL, "Cliente"
          LblGuiaR = Format$(Val(TxtGuiaRem), "00000000")
          CentrarFrame FrmGuiaRemision
          'FrmGuiaRemision.Top = 740
          FrmGuiaRemision.Visible = True
          MBoxFechaGRE.SetFocus
     Case "Suscripcion"
          RatonReloj
          Factura_No = TextFacturaNo.Text
          FSuscripcion.Show 1
     Case "Reserva"
          If DatInv.Por_Reservas Then
             'FrmReservas.Top = 740
             FrmReservas.Visible = True
             MBFechaIn.SetFocus
          Else
             FrmReservas.Visible = False
          End If
          'DCArticulo.width = AnchoDetalle
     Case "CopiarFactura"
          sSQL = "SELECT Factura " _
               & "FROM Facturas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Serie = '" & FA.Serie & "' " _
               & "ORDER BY Factura DESC "
          SelectDB_Combo DCFACopia, AdoFactura, sSQL, "Factura"
          FrmCopyFA.Top = MBoxFechaV.Top
          FrmCopyFA.Left = MBoxFechaV.Left
          FrmCopyFA.Caption = "SERIE: " & FA.Serie
          FrmCopyFA.Visible = True
          DCFACopia.SetFocus
   End Select
End Sub

Private Sub TextCant_Change()
  If IsNumeric(TextCant) Then
     If Val(TextCant) <> 0 And Val(TextVUnit) <> 0 Then Real1 = Redondear(CCur(TextCant) * CCur(TextVUnit), 2)
  Else
     Real1 = 0
  End If
  LabelVTotal.Caption = Format$(Real1, "#,##0.00")
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If AltDown And KeyCode = vbKeyP Then
     FrmPVP.Top = Label19.Top - 80
     FrmPVP.Left = Label19.Left - 10
     FrmPVP.Visible = True
     LstPVP.SetFocus
  End If
  
  PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
  If TextCant = "" Then TextCant = "0"
  If TextVUnit = "" Then TextVUnit = "0"
  CantAnterior = CCur(TextCant)
  Real1 = Redondear(CCur(TextCant) * CCur(TextVUnit), 2)
  TextVUnit = Format$(CCur(TextVUnit), "#,##0." & String(Dec_PVP, "0"))
  LabelVTotal.Caption = Format$(Real1, "#,##0.00")
End Sub

Private Sub TextComEjec_GotFocus()
  MarcarTexto TextComEjec
End Sub

Private Sub TextComEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyF11 Then
     DatInv.Reg_Sanitario = Ninguno
     DatInv.Procedencia = Ninguno
     DatInv.Modelo = Ninguno
     'DatInv.Serie_No = Ninguno
     DatInv.Fecha_Exp = FechaSistema
     DatInv.Fecha_Fab = FechaSistema
     Listar_Ordenes
     CentrarFrame FrmOrdenNo
     FrmOrdenNo.Visible = True
     LstOrden.SetFocus
     Ln_No_O = 0
  End If
  If CtrlDown And KeyCode = vbKeyF12 Then
     DatInv.Reg_Sanitario = Ninguno
     DatInv.Procedencia = Ninguno
     DatInv.Modelo = Ninguno
     'DatInv.Serie_No = Ninguno
     DatInv.Fecha_Exp = FechaSistema
     DatInv.Fecha_Fab = FechaSistema
     Listar_Lotes
     'Listar_Ordenes
     CentrarFrame FrmOrdenNo
     FrmOrdenNo.Visible = True
     LstOrden.SetFocus
     Ln_No_O = 0
  End If
End Sub

Private Sub TextComEjec_LostFocus()
    If Len(TextComEjec) > 1 Then
       FrmFechaV.Visible = True
       MBFechaVGR.SetFocus
    End If
    
    If DatInv.Codigo_Inv = "99.41" Then
       Titulo = "Formulario de Asignacion"
       Mensajes = "Ingresar Reembolso de Gastos con IVA?"
       If BoxMensaje = vbYes Then BanIVA = True Else BanIVA = False
    End If
End Sub

Private Sub TextComision_GotFocus()
  MarcarTexto TextComision
End Sub

Private Sub TextComision_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextComision_LostFocus()
  TextoValido TextComision, True
  If Val(TextComision) <= 0 Then
     MsgBox "El porcentaje debe ser mayor que cero"
     DCEjecutivo.SetFocus
  End If
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub


Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
  TextoValidoVar TextDesc, True
  Total_Desc = 0
  Si_No = False
  If MidStrg(TextDesc.Text, Len(TextDesc.Text), 1) = "%" Then
     Total_ME = Val(CCur(MidStrg(TextDesc.Text, 1, Len(TextDesc.Text) - 1)))
     Total_Desc = (Total_Con_IVA + Total_Sin_IVA) * (Total_ME / 100)
     Si_No = True
  Else
     Total_Desc = Val(CCur(TextDesc.Text))
  End If
  TextDesc = Redondear(Total_Desc, 2)
  TextoValido TextDesc, True
  Calculos_Totales_Factura FA
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
  Sec_Public = False
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
 If TextFacturaNo = "" Then TextFacturaNo = "0"
 'MBoxFecha.SetFocus
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
''     SelectDB_List DBLFact, AdoListFact, sSQL, "Factura"
''  End If
End Sub

Private Sub TextObs_LostFocus()
  TextoValido TextObs
End Sub

Private Sub TextVUnit_Change()
   LabelVTotal.Caption = Format$(Real1, "#,##0.00")
End Sub

Private Sub TextVUnit_GotFocus()
   MarcarTexto TextVUnit
   Valor_UnitA = Val(TextVUnit)
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
Dim SubTotal As Currency
Dim SubTotalServicio As Currency
Dim SubTotalIVA As Currency
Dim SubTotalPorcComision As Currency
Dim SubTotalDescuento As Currency

   If Not Mod_PVP Then TextVUnit = Valor_UnitA
   If DatInv.Serie_No = "" Then DatInv.Serie_No = Ninguno
  'MsgBox TipoFactura & vbCrLf & BanIVA
   Factura_No = Val(TextFacturaNo)
   TextoValido TextVUnit, True, , Dec_PVP
   TextoValido TextCant, True
  'TextoValido TextDesc1, True
   SubTotal = 0: SubTotalDescuento = 0: SubTotalIVA = 0: SubTotalPorcComision = 0
   NumMeses = 0: VUnitTemp = 0: Interes = 0: SubTotalServicio = 0
   With AdoAsientoF.Recordset
    If .RecordCount <= Cant_Item_FA Then
        If TxtDetalle <> Ninguno Then Producto = TxtDetalle
        TxtDetalle.Visible = False
       'Porcentaje por ejecutivo
        If Val(TextComision) > 0 Then SubTotalPorcComision = Redondear(Val(TextComision) / 100, 2)
       'SubTotal por producto
        SubTotal = Redondear(CCur(TextCant) * CDbl(TextVUnit), 2)
        If VUnitTemp > 0 Then SubTotal = Redondear(VUnitTemp, 2)
       'Descuento
        SubTotalDescuento = Redondear(SubTotal * (Redondear(Val(CDesc1.Text), 3) / 100), 2)
       'IVA = SubTotal - Descuento
       'MsgBox Porc_IVA
        If BanIVA And FA.TC <> "NV" Then SubTotalIVA = Redondear((SubTotal - SubTotalDescuento) * Porc_IVA, 4)
       'If TipoFactura = "OP" Then SubTotalIVA = 0
        If CDbl(TextVUnit) = 0 Then SubTotalIVA = 0
        LabelVTotal.Caption = Format$(SubTotal, "#,##0.00")
        If Porc_Serv > 0 Then SubTotalServicio = Redondear((SubTotal - SubTotalDescuento) * Porc_Serv, 2)
'''        If CheqCom.value = 1 Then FComision.Show 1
       'MsgBox Redondear(CDbl(TextVUnit), Dec_PVP) & " ..." & Redondear(Val(TextVUnit), Dec_PVP)
        If Len(Codigos) > 1 Then
           
           SetAdoAddNew "Asiento_F"
           SetAdoFields "CODIGO", Codigos
           SetAdoFields "CODIGO_L", CodigoL
           SetAdoFields "PRODUCTO", Producto
           SetAdoFields "REP", 0
           SetAdoFields "CANT", CCur(TextCant)
           SetAdoFields "PRECIO", Redondear(CDbl(TextVUnit), Dec_PVP)
           SetAdoFields "TOTAL", SubTotal
           SetAdoFields "VALOR_TOTAL", SubTotal - SubTotalDescuento + SubTotalIVA
           SetAdoFields "Total_Desc", SubTotalDescuento
           SetAdoFields "Total_IVA", SubTotalIVA
           SetAdoFields "SERVICIO", SubTotalServicio
           SetAdoFields "Cta", DatInv.Cta_Ventas
           SetAdoFields "Cta_SubMod", FA.SubCta
           SetAdoFields "CodBod", Cod_Bodega
           SetAdoFields "CodMar", Cod_Marca
           SetAdoFields "COD_BAR", CodigoInv1
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoFields "CORTE", VUnitTemp
           SetAdoFields "A_No", CByte(Ln_No)
           SetAdoFields "Fecha_V", MBFechaVGR
           SetAdoFields "Cod_Ejec", FA.Cod_Ejec
           SetAdoFields "Porc_C", SubTotalPorcComision
           SetAdoFields "Serie_No", DatInv.Serie_No
           If Len(TextComEjec) > 1 Then SetAdoFields "RUTA", TextComEjec
           If DatInv.Por_Reservas Then
              SetAdoFields "Fecha_IN", MBFechaIn
              SetAdoFields "Fecha_OUT", MBFechaOut
              SetAdoFields "Cant_Hab", TxtCantRooms
              SetAdoFields "Tipo_Hab", TxtTipoRooms
           End If
           If Len(Lote_No) > 1 Then   ' And Len(DatInv.Reg_Sanitario) > 1
              SetAdoFields "Lote_No", Lote_No
              SetAdoFields "Fecha_Fab", DatInv.Fecha_Fab
              SetAdoFields "Fecha_Exp", DatInv.Fecha_Exp
              SetAdoFields "Reg_Sanitario", DatInv.Reg_Sanitario
              SetAdoFields "Procedencia", DatInv.Procedencia
              SetAdoFields "Modelo", DatInv.Modelo
              SetAdoFields "SP", Sec_Public
           End If
           SetAdoFields "COSTO", DatInv.Costo
           If DatInv.Costo > 0 Then
              SetAdoFields "Cta_Inv", DatInv.Cta_Inventario
              SetAdoFields "Cta_Costo", DatInv.Cta_Costo_Venta
              SetAdoFields "Utilidad", Redondear(SubTotal - (CCur(TextCant) * DatInv.Costo), 2)
           End If
           If Len(No_Hab) > 1 Then SetAdoFields "HABIT", No_Hab
           
           SetAdoUpdate
           Ln_No = Ln_No + 1
        Else
           MsgBox "No ha seleccionado el codigo correcto, vuelva a ingresar"
        End If
    Else
        MsgBox "Ya no se puede ingresar más datos."
        Command1.SetFocus
    End If
   End With
   sSQL = "SELECT * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "ORDER BY A_No "
   SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|Total_IVA 4|."
   Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
   Calculos_Totales_Factura FA
   LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
   LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
   TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
   LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
   LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
   LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
   LblUtilidad.Caption = Format$(FA.Utilidad, "#,##0.00")
   DGAsientoF.Visible = True
   TextCant.Text = ""
   LabelVTotal.Caption = ""
   MarcarTexto TextDesc
   If (Redondear(CDbl(TextVUnit), Dec_PVP) < DatInv.Costo) And (DatInv.Costo > 0 And Len(DatInv.Cta_Inventario) > 3) Then
      MsgBox "Usted esta vendiendo por debajo del Costo de Produccion"
   End If
   DCArticulo.SetFocus
End Sub

Private Sub TxtCantRooms_GotFocus()
  MarcarTexto TxtCantRooms
End Sub

Private Sub TxtCantRooms_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCantRooms_LostFocus()
  TextoValido TxtCantRooms, True, , 2
End Sub

Private Sub TxtCompra_GotFocus()
  MarcarTexto TxtCompra
End Sub

Private Sub TxtCompra_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCompra_LostFocus()
  TextoValido TxtCompra, True, , 0
  If Val(TxtCompra) > 0 Then TxtCompra = Format(Val(TxtCompra), "0000000000")
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTextoFinal TxtDetalle
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then TxtDetalle.Visible = False
  If KeyCode = vbKeyTab Then
     TxtDetalle.Visible = False
     SiguienteControl
  End If
End Sub

Private Sub TxtDetalle_LostFocus()
  TxtDetalle.Visible = False
  'DCArticulo.width = AnchoDetalle
End Sub

Public Sub Insertar_Pedidos()
   If Len(Cta) > 1 And Len(Codigo) > 1 Then
      SetAdoAddNew "Asiento_F"
      SetAdoFields "CODIGO", Codigo
      SetAdoFields "CODIGO_L", CodigoL
      SetAdoFields "PRODUCTO", Codigo1
      SetAdoFields "CANT", Cantidad
      SetAdoFields "HABIT", Habitacion_No
      SetAdoFields "PRECIO", Precio
      SetAdoFields "TOTAL", Total
      SetAdoFields "Total_IVA", Total_IVA
      SetAdoFields "Cta", Cta
      SetAdoFields "Item", NumEmpresa
      SetAdoFields "CodigoU", CodigoUsuario
      SetAdoFields "Cta_SubMod", FA.SubCta
      SetAdoFields "RUTA", TextComEjec
      SetAdoFields "CodBod", Cod_Bodega
      SetAdoFields "CodMar", Cod_Marca
      SetAdoFields "A_No", CByte(Ln_No)
      If Val(TextComEjec.Text) > 0 Then
         SetAdoFields "Cod_Ejec", CodigoEjec
         SetAdoFields "Porc_C", Redondear(Val(TextComEjec) / 100, 4)
      End If
      SetAdoUpdate
      Ln_No = Ln_No + 1
   End If
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  TextoValido TxtEmail
  TxtEmail = LCase(TxtEmail)
  Actualiza_Email TxtEmail, FA.CodigoC
End Sub

Public Sub Llenar_Orden(LstOrdenP As ListBox)
Dim CantOrdenes As Byte
Dim IdxOrden As Integer
Dim OrdenP As Long
 DGAsientoF.Visible = False
 Select Case MidStrg(LstOrdenP.Text, 1, 4)
   Case "Lote"
        DatInv.Fecha_Exp = FechaSistema
        DatInv.Fecha_Fab = FechaSistema
        DatInv.Reg_Sanitario = Ninguno
        DatInv.Modelo = Ninguno
       'DatInv.Serie_No = Ninguno
        
        StockLote = 0
        For IdxOrden = 0 To LstOrdenP.ListCount - 1
          If LstOrdenP.Selected(IdxOrden) Then
             Lote_No = SinEspaciosDer(LstOrdenP.List(IdxOrden))
             sSQL = "SELECT Lote_No, Fecha_Fab, Fecha_Exp, Procedencia, Modelo, Serie_No, SUM(Entrada-Salida) As TotStock " _
                  & "FROM Trans_Kardex " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND T <> 'A' " _
                  & "AND Lote_No = '" & Lote_No & "' " _
                  & "GROUP BY Lote_No, Fecha_Fab, Fecha_Exp, Procedencia, Modelo, Serie_No " _
                  & "HAVING SUM(Entrada-Salida) <> 0 " _
                  & "ORDER BY Lote_No, Fecha_Fab, Fecha_Exp, Procedencia, Modelo, Serie_No "
             Select_Adodc AdoAux, sSQL
             With AdoAux.Recordset
              If .RecordCount > 0 Then
                  DatInv.Procedencia = .fields("Procedencia")
                  DatInv.Modelo = .fields("Modelo")
                  DatInv.Serie_No = .fields("Serie_No")
                  DatInv.Fecha_Exp = .fields("Fecha_Exp")
                  DatInv.Fecha_Fab = .fields("Fecha_Fab")
                 'DatInv.Reg_Sanitario = .Fields("Reg_Sanitario")
                  StockLote = .fields("TotStock")
                  IdxOrden = LstOrdenP.ListCount
              End If
             End With
          End If
        Next IdxOrden
   Case "Orde"
        If AdoOrden.Recordset.RecordCount > 0 Then
           CantOrdenes = 0
           For IdxOrden = 0 To LstOrdenP.ListCount - 1
               If LstOrdenP.Selected(IdxOrden) Then
                  Cadena = LstOrdenP.List(IdxOrden)
                  Cadena = MidStrg(Cadena, 11, Len(Cadena))
                  OrdenP = Val(SinEspaciosIzq(Cadena))
                 'MsgBox Cadena & vbCrLf & OrdenP
                  sSQL = "SELECT * " _
                       & "FROM Asiento_F " _
                       & "WHERE Item = '" & NumEmpresa & "' " _
                       & "AND CodigoU = '" & CodigoUsuario & "' " _
                       & "AND Numero = " & OrdenP & " "
                  Select_Adodc AdoAsientoF, sSQL
                  If AdoAsientoF.Recordset.RecordCount <= 0 Then
                     sSQL = "SELECT * " _
                          & "FROM Detalle_Factura " _
                          & "WHERE Item = '" & NumEmpresa & "' " _
                          & "AND Periodo = '" & Periodo_Contable & "' " _
                          & "AND T <> 'A' " _
                          & "AND TC = 'OP' " _
                          & "AND Factura = " & OrdenP & " " _
                          & "ORDER BY ID,Codigo "
                     Select_Adodc AdoAux, sSQL
                     RatonReloj
                     With AdoAux.Recordset
                      If .RecordCount > 0 Then
                          Do While Not .EOF
                             SetAdoAddNew "Asiento_F"
                             SetAdoFields "CODIGO", .fields("Codigo")
                             SetAdoFields "CODIGO_L", FA.Cod_CxC
                             SetAdoFields "PRODUCTO", .fields("Producto")
                             SetAdoFields "CANT", .fields("Cantidad")
                             SetAdoFields "PRECIO", .fields("Precio")
                             SetAdoFields "TOTAL", .fields("Total")
                             SetAdoFields "Total_Desc", .fields("Total_Desc")
                             SetAdoFields "Total_IVA", .fields("Total_IVA")
                             SetAdoFields "Serie_No", .fields("Serie_No")
                             SetAdoFields "CodBod", .fields("CodBodega")
                             SetAdoFields "COSTO", .fields("Costo")
                             SetAdoFields "Cta", FA.Cta_CxP
                             SetAdoFields "Item", NumEmpresa
                             SetAdoFields "CodigoU", CodigoUsuario
                             SetAdoFields "A_No", CByte(Ln_No)
                             SetAdoFields "Numero", OrdenP
                             SetAdoUpdate
                             Ln_No = Ln_No + 1
                            .MoveNext
                          Loop
                      End If
                     End With
                     CantOrdenes = CantOrdenes + 1
                  End If
               End If
           Next IdxOrden
        End If
        DGAsientoF.Visible = True
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        SQLDec = "PRECIO " & CStr(Dec_PVP) & "|CORTE " & CStr(Dec_PVP) & "|TOTAL 4|."
        Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
        Calculos_Totales_Factura FA
        LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
        LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
        TextDesc.Text = Format$(FA.Descuento, "#,##0.00")
        LabelServ.Caption = Format$(FA.Servicio, "#,##0.00")
        LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
        LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
        LblUtilidad.Caption = Format$(FA.Utilidad, "#,##0.00")
        TextObs.SetFocus
 End Select
 FrmOrdenNo.Visible = False
 RatonNormal
End Sub

Private Sub TxtLugarEntrega_GotFocus()
   MarcarTexto TxtLugarEntrega
End Sub

Private Sub TxtLugarEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLugarEntrega_LostFocus()
   TextoValido TxtLugarEntrega, , True
End Sub

Private Sub TxtNoches_GotFocus()
   MarcarTexto TxtNoches
End Sub

Private Sub TxtNoches_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtNoches_LostFocus()
   TextoValido TxtNoches, True, , 2
End Sub

Private Sub TxtPedido_GotFocus()
  MarcarTexto TxtPedido
End Sub

Private Sub TxtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPlaca_GotFocus()
  MarcarTexto TxtPlaca
End Sub

Private Sub TxtPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPlaca_LostFocus()
  TextoValido TxtPlaca, , True
End Sub

Public Sub Listar_Tipo_Beneficiarios(Grupo As String)
    RatonReloj
    DCCliente.Visible = False
    sSQL = "SELECT TOP 50 Cliente, CI_RUC, Codigo, Cta_CxP, Grupo, Cod_Ejec " _
         & "FROM Clientes " _
         & "WHERE FA <> " & Val(adFalse) & " " _
         & "AND T = 'N' "
    If Grupo <> Ninguno Then sSQL = sSQL & "AND Grupo = '" & Grupo & "' "
    sSQL = sSQL & "ORDER BY Cliente "
    SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
   'MsgBox "Desktop Test: " & Grupo & vbCrLf & AdoCliente.Recordset.RecordCount
    RatonNormal
    DCCliente.Visible = True
    DCCliente.SetFocus
End Sub

Public Sub Listar_Ordenes()
   LstOrden.Clear
   sSQL = "SELECT OP.Factura,OP.CodigoC,OP.Fecha,C.Cliente,C.Grupo,C.CI_RUC,C.TD " _
        & "FROM Facturas As OP, Clientes As C " _
        & "WHERE OP.Item = '" & NumEmpresa & "' " _
        & "AND OP.Periodo = '" & Periodo_Contable & "' " _
        & "AND OP.TC = 'OP' " _
        & "AND OP.T <> 'A' " _
        & "AND OP.CodigoC = C.Codigo " _
        & "ORDER BY OP.Factura "
   Select_Adodc AdoOrden, sSQL
   With AdoOrden.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           LstOrden.AddItem "Orden No. " & Format(.fields("Factura"), "000000000") & " - " & .fields("Cliente")
          .MoveNext
        Loop
        CentrarFrame FrmOrdenNo
        FrmOrdenNo.Visible = True
        LstOrden.SetFocus
    Else
        MsgBox "No existe Ordenes para procesar"
    End If
   End With
End Sub

Private Sub TxtTipoRooms_GotFocus()
  MarcarTexto TxtTipoRooms
End Sub

Private Sub TxtTipoRooms_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTipoRooms_LostFocus()
  TextoValido TxtTipoRooms, , True
End Sub

Private Sub TxtZona_GotFocus()
   MarcarTexto TxtZona
End Sub

Public Sub Listar_Lotes()
   LstOrden.Clear
   sSQL = "SELECT Lote_No " _
        & "FROM Trans_Kardex " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND T <> 'A' " _
        & "AND LEN(Lote_No) > 1 " _
        & "GROUP BY Lote_No " _
        & "ORDER BY Lote_No "
   Select_Adodc AdoAux, sSQL
   Select_Adodc AdoAux, sSQL
   With AdoAux.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           LstOrden.AddItem "Lote No. " & .fields("Lote_No")
          .MoveNext
        Loop
        LstOrden.Text = LstOrden.List(0)
    End If
   End With
End Sub

Private Sub TxtZona_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
