VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form IngLinea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " CUENTAS POR COBRAR"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView TVCatalogo 
      Height          =   3165
      Left            =   105
      TabIndex        =   57
      Top             =   315
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   5583
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Vencimiento de Facturas"
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
      Height          =   855
      Left            =   6615
      Picture         =   "IngLinea.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   1050
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   105
      TabIndex        =   5
      Top             =   3990
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "DATOS DE PROCESOS"
      TabPicture(0)   =   "IngLinea.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label22"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "MBoxCta_Anio_Anterior"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "MBoxCta_Venta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MBoxCta"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtItems"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CTipo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtPosY"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtPosFact"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtEspa"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtAncho"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtLargo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtNumFact"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtLogoFact"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CheqCtaVenta"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CheqMes"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "DATOS DEL S.R.I."
      TabPicture(1)   =   "IngLinea.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEstablecimiento"
      Tab(1).Control(1)=   "TxtNumSerieDos"
      Tab(1).Control(2)=   "TxtNumSerieUno"
      Tab(1).Control(3)=   "TxtNumSerietres1"
      Tab(1).Control(4)=   "TxtNumAutor"
      Tab(1).Control(5)=   "MBFechaVenc"
      Tab(1).Control(6)=   "MBFechaIni"
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(9)=   "Label17"
      Tab(1).Control(10)=   "Label14"
      Tab(1).Control(11)=   "Label15"
      Tab(1).Control(12)=   "Label18"
      Tab(1).ControlCount=   13
      Begin VB.Frame FrmEstablecimiento 
         Caption         =   "DATOS DEL ESTABLECIMIENTO:"
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
         Height          =   1590
         Left            =   -74895
         TabIndex        =   45
         Top             =   2100
         Visible         =   0   'False
         Width           =   7575
         Begin VB.TextBox TxtTelefonoEstab 
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
            MaxLength       =   10
            TabIndex        =   51
            Top             =   1155
            Width           =   2115
         End
         Begin VB.TextBox TxtLogoTipoEstab 
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
            MaxLength       =   10
            TabIndex        =   53
            Top             =   1155
            Width           =   2115
         End
         Begin VB.TextBox TxtDireccionEstab 
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
            MaxLength       =   60
            TabIndex        =   49
            Top             =   840
            Width           =   6105
         End
         Begin VB.TextBox TxtNombreEstab 
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
            MaxLength       =   60
            TabIndex        =   47
            Top             =   525
            Width           =   7365
         End
         Begin VB.Label Label23 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LOGOTIPO (GIF):"
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
            Left            =   3465
            TabIndex        =   52
            Top             =   1155
            Width           =   1905
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TELEFONO:"
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
            TabIndex        =   50
            Top             =   1155
            Width           =   1275
         End
         Begin VB.Label Label25 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIRECCION:"
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
            TabIndex        =   48
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label Label24 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NOMBRE DEL ESTABLECIMIENTO"
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
            Top             =   210
            Width           =   7365
         End
      End
      Begin VB.CheckBox CheqMes 
         Caption         =   "Facturacion por Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   12
         Top             =   1260
         Width           =   4215
      End
      Begin VB.TextBox TxtNumSerieDos 
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
         Left            =   -67965
         MaxLength       =   3
         TabIndex        =   44
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
         Top             =   1680
         Width           =   645
      End
      Begin VB.CheckBox CheqCtaVenta 
         Caption         =   "Cuenta de Venta si manejanos por sector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   10
         Top             =   840
         Width           =   5370
      End
      Begin VB.TextBox TxtNumSerieUno 
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
         Left            =   -68595
         MaxLength       =   3
         TabIndex        =   43
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
         Top             =   1680
         Width           =   645
      End
      Begin VB.TextBox TxtNumSerietres1 
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
         Left            =   -68490
         MaxLength       =   9
         TabIndex        =   37
         Text            =   "0000001"
         Top             =   840
         Width           =   1170
      End
      Begin VB.TextBox TxtNumAutor 
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
         Height          =   336
         Left            =   -68805
         MaxLength       =   37
         TabIndex        =   41
         Text            =   "0000000001"
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox TxtLogoFact 
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
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2100
         Width           =   2115
      End
      Begin VB.TextBox TxtNumFact 
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
         Left            =   3675
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1680
         Width           =   750
      End
      Begin VB.TextBox TxtLargo 
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
         Left            =   4725
         MaxLength       =   5
         TabIndex        =   29
         Top             =   3360
         Width           =   750
      End
      Begin VB.TextBox TxtAncho 
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
         MaxLength       =   5
         TabIndex        =   32
         Top             =   3360
         Width           =   750
      End
      Begin VB.TextBox TxtEspa 
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
         MaxLength       =   5
         TabIndex        =   27
         Top             =   3360
         Width           =   750
      End
      Begin VB.TextBox TxtPosFact 
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
         MaxLength       =   5
         TabIndex        =   23
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox TxtPosY 
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
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2940
         Width           =   750
      End
      Begin VB.ComboBox CTipo 
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
         Left            =   6825
         TabIndex        =   14
         Text            =   "FA"
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox TxtItems 
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
         MaxLength       =   5
         TabIndex        =   18
         Top             =   1680
         Width           =   750
      End
      Begin MSMask.MaskEdBox MBoxCta 
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   420
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBFechaVenc 
         Height          =   330
         Left            =   -72165
         TabIndex        =   39
         Top             =   1260
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
      Begin MSMask.MaskEdBox MBFechaIni 
         Height          =   330
         Left            =   -72165
         TabIndex        =   35
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
      Begin MSMask.MaskEdBox MBoxCta_Venta 
         Height          =   330
         Left            =   5565
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBoxCta_Anio_Anterior 
         Height          =   330
         Left            =   5565
         TabIndex        =   9
         Top             =   420
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CxC Año Anterior"
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
         Left            =   3885
         TabIndex        =   8
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SERIE DE FACTURA/NOTA DE VENTA (ESTAB. Y PUNTO DE VENTA)"
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
         Left            =   -74895
         TabIndex        =   42
         Top             =   1680
         Width           =   6315
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DATOS DEL S.R.I DE LA FACTURA/NOTA DE VENTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -74895
         TabIndex        =   33
         Top             =   420
         Width           =   7575
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA DE INICIO"
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
         Left            =   -74895
         TabIndex        =   34
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SECUENCIAL DE INICIO"
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
         Left            =   -70800
         TabIndex        =   36
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA DE VENCIMIENTO"
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
         Left            =   -74895
         TabIndex        =   38
         Top             =   1260
         Width           =   2745
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " AUTORIZACION"
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
         Left            =   -70800
         TabIndex        =   40
         Top             =   1260
         Width           =   2010
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CxC Clientes"
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
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FORMATO GRAFICO DEL DOCUMENTO (EXTENSION: GIF)"
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
         TabIndex        =   19
         Top             =   2100
         Width           =   5475
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DE FACTURAS POR PAGINA"
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
         TabIndex        =   15
         Top             =   1680
         Width           =   3585
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LARGO"
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
         Left            =   3885
         TabIndex        =   28
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ESPACIO Y POSICION DE LA COPIA DE LA FACTURA/NOTA DE VENTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   105
         TabIndex        =   21
         Top             =   2520
         Width           =   7575
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DE DOCUMENTO"
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
         Left            =   4410
         TabIndex        =   13
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ESPACIO ENTRE LA FACTURA"
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
         TabIndex        =   26
         Top             =   3360
         Width           =   2955
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " POSICION X DE LA FACTURA"
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
         Top             =   2940
         Width           =   2955
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " POSICION Y DE LA FACTURA"
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
         Left            =   3885
         TabIndex        =   24
         Top             =   2940
         Width           =   3060
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ANCHO"
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
         TabIndex        =   31
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   5565
         TabIndex        =   30
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ITEMS POR FACTURA"
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
         Left            =   4515
         TabIndex        =   17
         Top             =   1680
         Width           =   2430
      End
   End
   Begin VB.TextBox TextLinea 
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
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3570
      Width           =   4110
   End
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   210
      Top             =   945
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
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   210
      Top             =   630
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   210
      Top             =   1260
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
      Caption         =   "Art"
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
   Begin VB.TextBox TextCodigo 
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
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3570
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Left            =   6615
      Picture         =   "IngLinea.frx":047A
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   105
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6615
      Picture         =   "IngLinea.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1995
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   330
      Left            =   210
      Top             =   1575
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6825
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IngLinea.frx":0BC6
            Key             =   "Autorizaciones"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IngLinea.frx":14A0
            Key             =   "Autorizacion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IngLinea.frx":17BA
            Key             =   "Serie"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IngLinea.frx":1C0C
            Key             =   "Codigo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IngLinea.frx":24E6
            Key             =   "Otros"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DE LA CUENTA POR COBRAR"
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
      TabIndex        =   0
      Top             =   105
      Width           =   6420
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      TabIndex        =   1
      Top             =   3570
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DESCRIPCION"
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
      Left            =   2310
      TabIndex        =   3
      Top             =   3570
      Width           =   1485
   End
End
Attribute VB_Name = "IngLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VMBFechaIni As String
Dim VMBFechaVenc As String
Dim VTxtNumSerietres1  As String
Dim VTxtNumAutor As String
Dim VTxtNumSerieUno As String
Dim VTxtNumSerieDos As String
Dim Cta_Ini As String
Dim Cta_Fin As String

Private Sub CheqCtaVenta_Click()
  If CheqCtaVenta.value <> 0 Then MBoxCta_Venta.Visible = True Else MBoxCta_Venta.Visible = False
End Sub

Private Sub Command1_Click()
'''' If Nuevo Then GrabarCta (True) Else GrabarCta (False)
  GrabarArticulos
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload IngLinea
End Sub

Private Sub Command3_Click()
  sSQL = "SELECT Autorizacion,Serie,TC,MAX(Vencimiento) As Fecha_Venc " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY Autorizacion,Serie,TC " _
       & "ORDER BY Autorizacion,Serie,TC "
  Select_Adodc AdoArticulo, sSQL
  Cadena = ""
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          sSQL = "UPDATE Facturas " _
               & "SET Vencimiento = #" & BuscarFecha(.fields("Fecha_Venc")) & "# " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Autorizacion = '" & .fields("Autorizacion") & "' " _
               & "AND Serie = '" & .fields("Serie") & "' " _
               & "AND TC = '" & .fields("TC") & "' "
          Ejecutar_SQL_SP sSQL
         .MoveNext
       Loop
   End If
  End With
  MsgBox "Proceos Terminado, Proceda a verificar en Listar/Anular Facturas"
End Sub

Private Sub CTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
    RatonReloj
    sSQL = "UPDATE Catalogo_Lineas " _
         & "SET CxC_Anterior = CxC " _
         & "WHERE CxC_Anterior IN ('.','0') "
    Ejecutar_SQL_SP sSQL
    Llenar_CxC
    CTipo.Clear
    CTipo.AddItem "FA"
    CTipo.AddItem "FR"
    CTipo.AddItem "NV"
    CTipo.AddItem "PV"
    CTipo.AddItem "FT"
    CTipo.AddItem "NC"
    CTipo.AddItem "LC"
    CTipo.AddItem "GR"
    CTipo.AddItem "CP"
    CTipo.AddItem "OP"
    CTipo.AddItem "NDU"
    CTipo.AddItem "NDO"
    CTipo.AddItem "NPA"
    CTipo.AddItem "DES"
    
    CTipo.Text = "FA"
    Codigo = Ninguno
    FormatoMaskCta MBoxCta
    FormatoMaskCta MBoxCta_Venta
    FormatoMaskCta MBoxCta_Anio_Anterior
    LlenarArticulos Codigo
    RatonNormal
    TVCatalogo.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm IngLinea
   ConectarAdodc AdoArt
   ConectarAdodc AdoTipo
   ConectarAdodc AdoLinea
   ConectarAdodc AdoArticulo
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SubInd As Integer
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyI Then Cta_Ini = SinEspaciosIzq(TVCatalogo.SelectedItem)
  If CtrlDown And KeyCode = vbKeyU Then Cta_Fin = SinEspaciosIzq(TVCatalogo.SelectedItem)
  If CtrlDown And KeyCode = vbKeyDelete Then
     Sub_Cuenta = ""
     Mensajes = "Esta seguro de Eliminar la Cuenta: " & vbCrLf _
              & TVCatalogo.SelectedItem & "."
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then
        Cuenta = TVCatalogo.SelectedItem
        Codigo = TVCatalogo.SelectedItem.key
        Cta_Sup = CodigoCuentaSup(Codigo)
        Codigo = MidStrg(Codigo, Len(Cta_Sup) + 2, Len(Codigo))
        For SubInd = 1 To TVCatalogo.Nodes.Count
            If Cta_Sup = TVCatalogo.Nodes.Item(SubInd).key Then
               Sub_Cuenta = TVCatalogo.Nodes.Item(SubInd).Text
               SubInd = TVCatalogo.Nodes.Count
            End If
        Next SubInd
        TVCatalogo.Nodes.Remove TVCatalogo.SelectedItem.index
        sSQL = "DELETE * " _
             & "FROM Catalogo_Lineas " _
             & "WHERE Codigo = '" & Codigo & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND TL <> " & Val(adFalse) & " "
        Ejecutar_SQL_SP sSQL
        'Llenar_CxC
        TVCatalogo.SelectedItem = Sub_Cuenta
     End If
  End If

End Sub

Private Sub TVCatalogo_LostFocus()
  Cuenta = TVCatalogo.SelectedItem
  Codigo = TVCatalogo.SelectedItem.key
  Cta_Sup = CodigoCuentaSup(Codigo)
  Sub_Cuenta = TVCatalogo.SelectedItem
  Codigo = MidStrg(Codigo, Len(Cta_Sup) + 2, Len(Codigo))
 'MsgBox Codigo & vbCrLf & Cta_Sup
  LlenarArticulos Codigo
End Sub

Private Sub MBFechaIni_GotFocus()
  MarcarTexto MBFechaIni
End Sub

Private Sub MBFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaIni_LostFocus()
  FechaValida MBFechaIni
  VMBFechaIni = MBFechaIni
  VMBFechaVenc = SiguienteAnio(MBFechaIni)
  MBFechaVenc = VMBFechaVenc
End Sub

Private Sub MBFechaVenc_GotFocus()
  MarcarTexto MBFechaVenc
End Sub

Private Sub MBFechaVenc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaVenc_LostFocus()
  FechaValida MBFechaIni
  FechaValida MBFechaVenc
  If CFechaLong(MBFechaVenc) > CFechaLong(VMBFechaVenc) Then
     MsgBox "La fecha de vencimiento se encuentra fuere del rango"
     MBFechaVenc.SetFocus
  End If
End Sub

Private Sub MBoxCta_GotFocus()
  MarcarTexto MBoxCta
End Sub

Private Sub MBoxCta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCodigo_GotFocus()
  MarcarTexto TextCodigo
End Sub

Private Sub TextCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, , True
End Sub

Private Sub TextLinea_GotFocus()
  MarcarTexto TextLinea
End Sub

Private Sub TextLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextLinea_LostFocus()
  TextoValido TextLinea
End Sub

Public Sub LlenarArticulos(CodigoArt As String)
  VMBFechaIni = Ninguno
  VMBFechaVenc = Ninguno
  VTxtNumSerietres1 = Ninguno
  VTxtNumAutor = Ninguno
  VTxtNumSerieUno = Ninguno
  VTxtNumSerieDos = Ninguno
  CheqMes.value = 0
  CheqCtaVenta.value = 0
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Codigo ='" & CodigoArt & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TL <> " & Val(adFalse) & " "
  Select_Adodc AdoLinea, sSQL
  With AdoLinea.Recordset
   If .RecordCount > 0 Then
       TextCodigo.Text = .fields("Codigo")
       TextLinea.Text = .fields("Concepto")
       MBoxCta.Text = FormatoCodigoCta(.fields("CxC"))
       If Len(.fields("CxC_Anterior")) = 1 Then
          MBoxCta_Anio_Anterior = FormatoCodigoCta(.fields("CxC"))
       Else
          MBoxCta_Anio_Anterior = FormatoCodigoCta(.fields("CxC_Anterior"))
       End If
       MBoxCta_Venta.Text = FormatoCodigoCta(.fields("Cta_Venta"))
       TxtLogoFact.Text = .fields("Logo_Factura")
       TxtLargo.Text = Format$(.fields("Largo"), "#0.00")
       TxtAncho.Text = Format$(.fields("Ancho"), "#0.00")
       TxtEspa.Text = Format$(.fields("Espacios"), "#0.00")
       TxtPosFact.Text = Format$(.fields("Pos_Factura"), "#0.00")
       TxtPosY.Text = Format$(.fields("Pos_Y_Fact"), "#0.00")
       TxtNumFact.Text = Format$(.fields("Fact_Pag"), "00")
       TxtItems.Text = Format$(.fields("ItemsxFA"), "#0.00")
       CTipo.Text = .fields("Fact")
      'SRI
       MBFechaIni = .fields("Fecha")
       MBFechaVenc = .fields("Vencimiento")
       TxtNumSerietres1 = Format$(.fields("Secuencial"), "000000000")
       TxtNumAutor = .fields("Autorizacion")
       TxtNumSerieUno = MidStrg(.fields("Serie"), 1, 3)
       TxtNumSerieDos = MidStrg(.fields("Serie"), 4, 3)
       TxtNombreEstab = .fields("Nombre_Establecimiento")
       TxtDireccionEstab = .fields("Direccion_Establecimiento")
       TxtTelefonoEstab = .fields("Telefono_Estab")
       TxtLogoTipoEstab = .fields("Logo_Tipo_Estab")
              
       VMBFechaIni = .fields("Fecha")
       VMBFechaVenc = .fields("Vencimiento")
       VTxtNumSerietres1 = Format$(.fields("Secuencial"), "000000000")
       VTxtNumAutor = .fields("Autorizacion")
       VTxtNumSerieUno = MidStrg(.fields("Serie"), 1, 3)
       VTxtNumSerieDos = MidStrg(.fields("Serie"), 4, 3)
       If .fields("Imp_Mes") Then
          MBoxCta_Venta.Visible = True
          CheqMes.value = 1
       Else
          MBoxCta_Venta.Visible = False
       End If
       If Len(.fields("Cta_Venta")) > 1 Then
          MBoxCta_Venta.Visible = True
          CheqCtaVenta.value = 1
       Else
          MBoxCta_Venta.Visible = False
       End If
       TextLinea.SetFocus
   Else
       TextCodigo.Text = ""
       TextLinea.Text = "NO PROCESABLE"
       MBoxCta.Text = LimpiarCtas
       TxtLogoFact.Text = ""
       TxtLargo.Text = "0.00"
       TxtAncho.Text = "0.00"
       TxtEspa.Text = "0.00"
       TxtPosFact.Text = "0.00"
       TxtPosY.Text = "0.00"
       TxtNumFact.Text = "00"
       TxtItems.Text = "0.00"
       CTipo.Text = ""
      'SRI
       MBFechaIni = FechaSistema
       MBFechaVenc = FechaSistema
       TxtNumSerietres1 = "000000000"
       TxtNumAutor = "000000000"
       TxtNumSerieUno = "000"
       TxtNumSerieDos = "000"
   End If
  End With
End Sub

Public Sub GrabarArticulos()
  Codigo = TextCodigo.Text
  TextoValido TxtLargo
  TextoValido TxtAncho
  TextoValido TxtEspa
  TextoValido TxtPosFact
  If CTipo.Text = "" Then CTipo.Text = "FA"
  If Val(TxtSerie) < 1001 Then TxtSerie = "001001"
  Mensajes = "Esta seguro de Grabar el Producto: " _
           & TextLinea.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then

     sSQL = "SELECT * " _
          & "FROM Catalogo_Lineas " _
          & "WHERE Codigo = '" & Codigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TL <> " & Val(adFalse) & " "
     Select_Adodc AdoLinea, sSQL, False
     With AdoLinea.Recordset
          IE = TVCatalogo.SelectedItem.index
          If .RecordCount <= 0 Then
              Control_Procesos "F", "Creación de Punto de Venta de " & CTipo & "-" & TxtNumSerieUno & TxtNumSerieDos
              Control_Procesos "F", "Creación de Fecha de Vencimiento de " & TextCodigo & " " & MBFechaVenc
              Control_Procesos "F", "Creación de Autorización de " & TextCodigo & " " & TxtNumAutor
              Control_Procesos "F", "Creación de Serie de " & TextCodigo & " " & TxtNumSerieUno & TxtNumSerieDos
              Control_Procesos "F", "Creación de Secuencial Inicial de " & TextCodigo & " " & TxtNumSerietres1
              
              SetAddNew AdoLinea
              SetFields AdoLinea, "Codigo", TextCodigo
              SetFields AdoLinea, "Item", NumEmpresa
              SetFields AdoLinea, "Periodo", Periodo_Contable
              SetFields AdoLinea, "TL", True
              Codigo = "A." & TxtNumAutor & "." & TxtNumSerieUno & TxtNumSerieDos & "." & CTipo & "." & TextCodigo
              Cuenta = TextLinea
          Else
              Control_Procesos "F", "Modificación de Punto de Venta de " & CTipo & "-" & TxtNumSerieUno & TxtNumSerieDos
              sSQL = "DELETE * " _
                   & "FROM Catalogo_Lineas " _
                   & "WHERE Codigo = '" & Codigo & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND TL <> " & Val(adFalse) & " "
              Ejecutar_SQL_SP sSQL
              sSQL = "SELECT * " _
                   & "FROM Catalogo_Lineas " _
                   & "WHERE Codigo = '" & Codigo & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND TL <> " & Val(adFalse) & " "
              Select_Adodc AdoLinea, sSQL, False
              If MBFechaVenc <> .fields("Vencimiento") Then Control_Procesos "F", "Modifico: Fecha de Vencimiento de " & TextCodigo & " " & MBFechaVenc
              If TxtNumAutor <> .fields("Autorizacion") Then Control_Procesos "F", "Modifico: Autorización de " & TextCodigo & " " & TxtNumAutor
              If TxtNumSerieUno & TxtNumSerieDos <> .fields("Serie") Then Control_Procesos "F", "Modifico: Serie de " & TextCodigo & " " & TxtNumSerieUno & TxtNumSerieDos
              If TxtNumSerietres1 <> .fields("Secuencial") Then Control_Procesos "F", "Modifico: Secuencial Inicial de " & TextCodigo & " " & TxtNumSerietres1
              SetAddNew AdoLinea
              SetFields AdoLinea, "Codigo", TextCodigo
              SetFields AdoLinea, "Item", NumEmpresa
              SetFields AdoLinea, "Periodo", Periodo_Contable
              SetFields AdoLinea, "TL", True
              TVCatalogo.Nodes(IE).Text = TextLinea
          End If
          SetFields AdoLinea, "Concepto", TextLinea
          SetFields AdoLinea, "CxC", CambioCodigoCta(MBoxCta)
          SetFields AdoLinea, "CxC_Anterior", CambioCodigoCta(MBoxCta_Anio_Anterior)
          SetFields AdoLinea, "Cta_Venta", CambioCodigoCta(MBoxCta_Venta)
          If CheqMes.value <> 0 Then SetFields AdoLinea, "Imp_Mes", True
          SetFields AdoLinea, "Logo_Factura", TxtLogoFact
          SetFields AdoLinea, "Largo", TxtLargo
          SetFields AdoLinea, "Ancho", TxtAncho
          SetFields AdoLinea, "Espacios", TxtEspa
          SetFields AdoLinea, "Pos_Factura", TxtPosFact
          SetFields AdoLinea, "Pos_Y_Fact", TxtPosY
          SetFields AdoLinea, "Fact_Pag", TxtNumFact
          SetFields AdoLinea, "ItemsxFA", TxtItems
          SetFields AdoLinea, "Fact", CTipo
         'SRI
          SetFields AdoLinea, "Fecha", MBFechaIni
          SetFields AdoLinea, "Vencimiento", MBFechaVenc
          SetFields AdoLinea, "Secuencial", TxtNumSerietres1
          SetFields AdoLinea, "Autorizacion", TxtNumAutor
          SetFields AdoLinea, "Serie", TxtNumSerieUno & TxtNumSerieDos
          SetFields AdoLinea, "Nombre_Establecimiento", TxtNombreEstab
          SetFields AdoLinea, "Direccion_Establecimiento", TxtDireccionEstab
          SetFields AdoLinea, "Telefono_Estab", TxtTelefonoEstab
          SetFields AdoLinea, "Logo_Tipo_Estab", TxtLogoTipoEstab
          SetUpdate AdoLinea
          TVCatalogo.Refresh
     End With
     
     sSQL = "SELECT * " _
          & "FROM Facturas_Formatos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cod_CxC = '" & Codigo & "' " _
          & "AND Serie = '" & TxtNumSerieUno & TxtNumSerieDos & "' " _
          & "AND Autorizacion = '" & TxtNumAutor & "' "
     Select_Adodc AdoArt, sSQL, False
     With AdoArt.Recordset
      If .RecordCount <= 0 Then
          SetAddNew AdoArt
          SetFields AdoArt, "Cod_CxC", TextCodigo
          SetFields AdoArt, "Item", NumEmpresa
          SetFields AdoArt, "Periodo", Periodo_Contable
      End If
      SetFields AdoArt, "Concepto", TextLinea
      SetFields AdoArt, "Formato_Factura", TxtLogoFact
      SetFields AdoArt, "Largo", TxtLargo
      SetFields AdoArt, "Ancho", TxtAncho
      SetFields AdoArt, "Espacios", TxtEspa
      SetFields AdoArt, "Pos_Factura", TxtPosFact
      SetFields AdoArt, "Pos_Y_Fact", TxtPosY
      SetFields AdoArt, "Fact_Pag", TxtNumFact
      SetFields AdoArt, "TC", CTipo
     'SRI
      SetFields AdoArt, "Fecha_Inicio", MBFechaIni
      SetFields AdoArt, "Fecha_Final", MBFechaVenc
      SetFields AdoArt, "Autorizacion", TxtNumAutor
      SetFields AdoArt, "Serie", TxtNumSerieUno & TxtNumSerieDos
      SetFields AdoArt, "Nombre_Establecimiento", TxtNombreEstab
      SetFields AdoArt, "Direccion_Establecimiento", TxtDireccionEstab
      SetFields AdoArt, "Telefono_Estab", TxtTelefonoEstab
      SetFields AdoArt, "Logo_Tipo_Estab", TxtLogoTipoEstab
      SetUpdate AdoArt
     End With
    'Numeracion de FA,NC,GR,ETC
     Select Case CTipo
       Case "NC"
            sSQL = "SELECT Periodo, Item, 'NC' As TC, Serie_NC As Serie_X, MAX(Secuencial_NC) As TC_No " _
                 & "FROM Trans_Abonos " _
                 & "WHERE Serie_NC = '" & TxtNumSerieUno & TxtNumSerieDos & "' " _
                 & "GROUP BY Periodo, Item, Serie_NC " _
                 & "ORDER BY Periodo, Item, Serie_NC "
       Case "GR"
            sSQL = "SELECT Periodo, Item, 'GR' As TC, Serie_GR As Serie_X, MAX(Remision) As TC_No " _
                 & "FROM Facturas_Auxiliares " _
                 & "WHERE Serie_GR = '" & TxtNumSerieUno & TxtNumSerieDos & "' " _
                 & "GROUP BY Periodo, Item, Serie_GR " _
                 & "ORDER BY Periodo, Item, Serie_GR "
       Case Else
            sSQL = "SELECT Periodo, Item, TC, Serie As Serie_X, MAX(Factura) As TC_No " _
                 & "FROM Facturas " _
                 & "WHERE TC = '" & CTipo & "' " _
                 & "AND Serie = '" & TxtNumSerieUno & TxtNumSerieDos & "' " _
                 & "GROUP BY Periodo, Item, TC, Serie " _
                 & "ORDER BY Periodo, Item, TC, Serie "
     End Select
     Select_Adodc AdoArticulo, sSQL, False
     With AdoArticulo.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             sSQL = "SELECT * " _
                  & "FROM Codigos " _
                  & "WHERE Periodo = '" & .fields("Periodo") & "' " _
                  & "AND Item = '" & .fields("Item") & "' " _
                  & "AND Concepto = '" & .fields("TC") & "_SERIE_" & .fields("Serie_X") & "' "
             Select_Adodc AdoArt, sSQL, False
             If AdoArt.Recordset.RecordCount > 0 Then
                AdoArt.Recordset.fields("Numero") = .fields("TC_No") + 1
                AdoArt.Recordset.Update
             Else
                SetAdoAddNew "Codigos"
                SetAdoFields "Item", .fields("Item")
                SetAdoFields "Periodo", .fields("Periodo")
                SetAdoFields "Concepto", .fields("TC") & "_SERIE_" & .fields("Serie_X")
                SetAdoFields "Numero", .fields("TC_No") + 1
                SetAdoUpdate
             End If
            .MoveNext
          Loop
      Else
          SetAdoAddNew "Codigos"
          SetAdoFields "Item", NumEmpresa
          SetAdoFields "Periodo", Periodo_Contable
          SetAdoFields "Concepto", CTipo & "_SERIE_" & TxtNumSerieUno & TxtNumSerieDos
          SetAdoFields "Numero", Val(TxtNumSerietres1)
          SetAdoUpdate
      End If
     End With
  End If
  Nuevo = False
  RatonNormal
  Llenar_CxC
  MsgBox "El proceso de Grabación se realizó con éxito"
End Sub

Private Sub TxtAncho_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDireccionEstab_GotFocus()
  MarcarTexto TxtDireccionEstab
End Sub

Private Sub TxtDireccionEstab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDireccionEstab_LostFocus()
  TextoValido TxtDireccionEstab
End Sub

Private Sub TxtEspa_GotFocus()
  MarcarTexto TxtEspa
End Sub

Private Sub TxtEspa_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtItems_GotFocus()
  MarcarTexto TxtItems
End Sub

Private Sub TxtItems_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtItems_LostFocus()
  TextoValido TxtItems, True
End Sub

Private Sub TxtLargo_GotFocus()
  MarcarTexto TxtLargo
End Sub

Private Sub TxtLargo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLargo_LostFocus()
  TextoValido TxtLargo, True
  If Val(TxtLargo.Text) <= 0 Then TxtLargo.Text = "15"
End Sub

Private Sub TxtAncho_GotFocus()
  MarcarTexto TxtAncho
End Sub

Private Sub TxtAncho_LostFocus()
  TextoValido TxtAncho, True
  If Val(TxtAncho.Text) <= 0 Then TxtAncho.Text = "19"
  SSTab1.Tab = 1
  MBFechaIni.SetFocus
End Sub

Private Sub TxtLogoFact_GotFocus()
  MarcarTexto TxtLogoFact
End Sub

Private Sub TxtLogoFact_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLogoFact_LostFocus()
  TextoValido TxtLogoFact, , True
  If TxtLogoFact.Text = Ninguno Then TxtLogoFact.Text = "NINGUNO"
End Sub

Private Sub TxtLogoTipoEstab_GotFocus()
  MarcarTexto TxtLogoTipoEstab
End Sub

Private Sub TxtLogoTipoEstab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLogoTipoEstab_LostFocus()
  TextoValido TxtLogoTipoEstab, , True
  FrmEstablecimiento.Visible = False
  SSTab1.Tab = 0
  MBoxCta.SetFocus
End Sub

Private Sub TxtNombreEstab_GotFocus()
  MarcarTexto TxtNombreEstab
End Sub

Private Sub TxtNombreEstab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNombreEstab_LostFocus()
  TextoValido TxtNombreEstab, , True
End Sub

Private Sub TxtNumAutor_GotFocus()
  MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
''  If Len(TxtNumAutor) <> 10 Then
''     MsgBox "La autorizacion debe tener 10 digitos"
''     TxtNumAutor.SetFocus
''  Else
''     TxtNumAutor = Format$(Val(TxtNumAutor), "0000000000")
''  End If
End Sub

Private Sub TxtNumFact_GotFocus()
  MarcarTexto TxtNumFact
End Sub

Private Sub TxtNumFact_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_GotFocus()
  MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
  If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
  TxtNumSerieDos = Format$(Val(TxtNumSerieDos), "000")
  If Val(TxtNumSerieUno) > 1 Then
     FrmEstablecimiento.Visible = True
     TxtNombreEstab.SetFocus
  Else
     FrmEstablecimiento.Visible = False
     SSTab1.Tab = 0
     MBoxCta.SetFocus
  End If
End Sub

Private Sub TxtNumSerietres1_GotFocus()
  MarcarTexto TxtNumSerietres1
End Sub

Private Sub TxtNumSerietres1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres1_LostFocus()
  If Val(TxtNumSerietres1) <= 0 Then TxtNumSerietres1 = "1"
  TxtNumSerietres1 = Format$(Val(TxtNumSerietres1), "000000000")
End Sub

Private Sub TxtNumSerieUno_GotFocus()
  MarcarTexto TxtNumSerieUno
End Sub

Private Sub TxtNumSerieUno_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUno_LostFocus()
  If Val(TxtNumSerieUno) <= 0 Then TxtNumSerieUno = "001"
  TxtNumSerieUno = Format$(Val(TxtNumSerieUno), "000")
End Sub

Private Sub TxtPosFact_GotFocus()
  MarcarTexto TxtPosFact
End Sub

Private Sub TxtPosFact_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPosY_GotFocus()
  MarcarTexto TxtPosY
End Sub

Private Sub TxtPosY_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub AddNewCta(TipoTC As String, Codigo As String, Detalle As String)
Dim SubInd As Integer
Dim inserteKey As Boolean
  If Len(Codigo) = 1 Then
     TVCatalogo.Nodes.Add , , Codigo, Detalle, ImageList1.ListImages(1).key, ImageList1.ListImages(1).key
     TVCatalogo.Tag = Codigo
  Else
     Select Case TipoTC
       Case "": IE = 1
       Case "A": IE = 2
       Case "S": IE = 3
       Case "T": IE = 4
       Case "D": IE = 5
     End Select
     Cta_Sup = CodigoCuentaSup(Codigo)
     inserteKey = True
     For SubInd = 1 To TVCatalogo.Nodes.Count
         If TVCatalogo.Nodes(SubInd).key = Codigo Then inserteKey = False
     Next SubInd
     If inserteKey Then
        TVCatalogo.Nodes.Add Cta_Sup, tvwChild, Codigo, Detalle, ImageList1.ListImages(IE).key, ImageList1.ListImages(IE).key
        TVCatalogo.Tag = Codigo ' MidStrg(Codigo, 2, Len(Codigo))
     End If
  End If
End Sub

Public Sub Llenar_CxC()
   RatonReloj
   TVCatalogo.Nodes.Clear
   Codigo = "A"
   Cuenta = "AUTORIZACIONES"
   AddNewCta "", Codigo, Cuenta
   sSQL = "SELECT * " _
        & "FROM Catalogo_Lineas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TL <> " & Val(adFalse) & " " _
        & "ORDER BY Autorizacion,Serie,Fact,Codigo "
   Select_Adodc AdoArt, sSQL
   With AdoArt.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           Codigo = "A." & .fields("Autorizacion")
           Cuenta = .fields("Autorizacion")
           AddNewCta "A", Codigo, Cuenta
           
           Codigo = "A." & .fields("Autorizacion") & "." & .fields("Serie")
           Cuenta = .fields("Serie")
           AddNewCta "S", Codigo, Cuenta
           
           Codigo = "A." & .fields("Autorizacion") & "." & .fields("Serie") & "." & .fields("Fact")
           Cuenta = .fields("Fact")
           AddNewCta "T", Codigo, Cuenta
           
           Codigo = "A." & .fields("Autorizacion") & "." & .fields("Serie") & "." & .fields("Fact") & "." & .fields("Codigo")
           Cuenta = .fields("Concepto")
           AddNewCta "D", Codigo, Cuenta
           
          .MoveNext
        Loop
    End If
   End With
   RatonNormal
End Sub

Private Sub TxtTelefonoEstab_GotFocus()
  MarcarTexto TxtTelefonoEstab
End Sub

Private Sub TxtTelefonoEstab_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoEstab_LostFocus()
  TxtTelefonoEstab = Format(Val(TxtTelefonoEstab), "000000000")
End Sub
