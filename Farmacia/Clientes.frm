VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Clientes"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Cliente"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Cliente"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Cliente"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Cliente"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior Cliente"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente Cliente"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultumo Cliente"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8835
      Begin MSDataListLib.DataCombo DCGrupoNo 
         DataSource      =   "AdoGrupoNo"
         Height          =   315
         Left            =   6240
         TabIndex        =   81
         Top             =   225
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataList DLCliente 
         DataSource      =   "AdoDBCliente"
         Height          =   840
         Left            =   2200
         TabIndex        =   80
         Top             =   600
         Width           =   5320
         _ExtentX        =   9366
         _ExtentY        =   1482
         _Version        =   393216
         BackColor       =   16777215
      End
      Begin VB.OptionButton OpcEjecutivo 
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
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   630
         Width           =   2010
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
         Height          =   960
         Left            =   7665
         Picture         =   "Clientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   315
         Width           =   1065
      End
      Begin VB.OptionButton OpcResp 
         Caption         =   "Responsable Area"
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
         TabIndex        =   5
         Top             =   945
         Width           =   2010
      End
      Begin VB.OptionButton OpcCliente 
         Caption         =   "Cliente"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE DEL CLIENTE             GRUPO No."
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
         Left            =   2205
         TabIndex        =   2
         Top             =   210
         Width           =   4005
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   11139
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&1.- DATOS PRINCIPALES"
      TabPicture(0)   =   "Clientes.frx":0282
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LabelCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CheqMult"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CheqContrato"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TextCliente"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TextEmpresa"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TextCiudad"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "MBoxComision"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TextEmail"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MBoxRUC"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "MBoxFAX"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "MBoxTelef1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TextDireccion"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "MBoxTelef2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TextDireccion1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TextGrupo"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Frame3"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "MBoxFecha_N"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TextProfesion"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "&2.- DATOS SECUNDARIOS"
      TabPicture(1)   =   "Clientes.frx":029E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TextPaisE"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "TextPlacas"
      Tab(1).Control(3)=   "TextTarj"
      Tab(1).Control(4)=   "TextTarjCred"
      Tab(1).Control(5)=   "TextNacionalidad"
      Tab(1).Control(6)=   "TextEstCiv"
      Tab(1).Control(7)=   "TextPais"
      Tab(1).Control(8)=   "MBoxRUCE"
      Tab(1).Control(9)=   "Label26"
      Tab(1).Control(10)=   "Label25"
      Tab(1).Control(11)=   "Label24"
      Tab(1).Control(12)=   "Label23"
      Tab(1).Control(13)=   "Label31"
      Tab(1).Control(14)=   "Label22"
      Tab(1).Control(15)=   "Label21"
      Tab(1).Control(16)=   "Label18"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "&3.- RESERVACIONES"
      TabPicture(2)   =   "Clientes.frx":02BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2"
      Tab(2).Control(1)=   "TextDias"
      Tab(2).Control(2)=   "TextCantHab"
      Tab(2).Control(3)=   "Label28"
      Tab(2).Control(4)=   "Label27"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command2 
         Caption         =   "&Reservacion de Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   -74895
         Picture         =   "Clientes.frx":02D6
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2625
         Width           =   1170
      End
      Begin VB.TextBox TextPaisE 
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
         Left            =   -69855
         MaxLength       =   15
         TabIndex        =   77
         Top             =   2835
         Width           =   2430
      End
      Begin VB.TextBox TextDias 
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
         MaxLength       =   3
         TabIndex        =   74
         Top             =   2205
         Width           =   1590
      End
      Begin VB.TextBox TextCantHab 
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
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   73
         Top             =   2205
         Width           =   1800
      End
      Begin VB.Frame Frame4 
         Caption         =   "CUAL EL MOTIVO DE SU VISITA ? "
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
         Left            =   -74895
         TabIndex        =   65
         Top             =   3885
         Width           =   4635
         Begin VB.OptionButton Option2 
            Caption         =   "PLEASURE / PLACER"
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
            TabIndex        =   67
            Top             =   630
            Width           =   2325
         End
         Begin VB.OptionButton Option1 
            Caption         =   "BUSINESS / NEGOCIOS"
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
            TabIndex        =   66
            Top             =   315
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.TextBox TextPlacas 
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
         MaxLength       =   15
         TabIndex        =   63
         Top             =   3465
         Width           =   3480
      End
      Begin VB.TextBox TextTarj 
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
         MaxLength       =   20
         TabIndex        =   62
         Top             =   2835
         Width           =   2010
      End
      Begin VB.TextBox TextTarjCred 
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
         Left            =   -72900
         MaxLength       =   25
         TabIndex        =   61
         Top             =   2835
         Width           =   2745
      End
      Begin VB.TextBox TextNacionalidad 
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
         Left            =   -69855
         MaxLength       =   15
         TabIndex        =   60
         Top             =   3465
         Width           =   2430
      End
      Begin VB.TextBox TextEstCiv 
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
         MaxLength       =   15
         TabIndex        =   57
         Top             =   2205
         Width           =   2640
      End
      Begin VB.TextBox TextPais 
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
         Left            =   -72270
         MaxLength       =   15
         TabIndex        =   56
         Top             =   2205
         Width           =   2325
      End
      Begin VB.TextBox TextProfesion 
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
         MaxLength       =   40
         TabIndex        =   35
         Top             =   4620
         Width           =   4110
      End
      Begin MSMask.MaskEdBox MBoxFecha_N 
         Height          =   330
         Left            =   105
         TabIndex        =   33
         Top             =   4620
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
      Begin VB.Frame Frame3 
         Height          =   540
         Left            =   6615
         TabIndex        =   23
         Top             =   3045
         Width           =   2220
         Begin VB.OptionButton OpcM 
            Caption         =   "Mujer"
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
            TabIndex        =   25
            Top             =   210
            Width           =   855
         End
         Begin VB.OptionButton OpcH 
            Caption         =   "Hombre"
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
            TabIndex        =   24
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.TextBox TextGrupo 
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
         Left            =   5460
         MaxLength       =   15
         TabIndex        =   37
         Top             =   4620
         Width           =   1065
      End
      Begin VB.TextBox TextDireccion1 
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
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4095
         Width           =   5160
      End
      Begin MSMask.MaskEdBox MBoxTelef2 
         Height          =   330
         Left            =   105
         TabIndex        =   28
         Top             =   4095
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
         Format          =   "CC-CCC-CCC"
         Mask            =   "CC-CCC-CCC"
         PromptChar      =   "0"
      End
      Begin VB.TextBox TextDireccion 
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
         MaxLength       =   50
         TabIndex        =   30
         Top             =   3780
         Width           =   5160
      End
      Begin MSMask.MaskEdBox MBoxTelef1 
         Height          =   330
         Left            =   105
         TabIndex        =   27
         Top             =   3780
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
         Format          =   "CC-CCC-CCC"
         Mask            =   "CC-CCC-CCC"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MBoxFAX 
         Height          =   330
         Left            =   5250
         TabIndex        =   22
         Top             =   3255
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
         Format          =   "CC-CCC-CCC"
         Mask            =   "CC-CCC-CCC"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MBoxRUC 
         Height          =   330
         Left            =   3465
         TabIndex        =   20
         Top             =   3255
         Width           =   1800
         _ExtentX        =   3175
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
         Format          =   "CCCCCCCCC-C-CCC"
         Mask            =   "CCCCCCCCC-C-CCC"
         PromptChar      =   "0"
      End
      Begin VB.TextBox TextEmail 
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
         TabIndex        =   18
         Top             =   3255
         Width           =   3375
      End
      Begin MSMask.MaskEdBox MBoxComision 
         Height          =   330
         Left            =   5565
         TabIndex        =   15
         Top             =   2730
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "00.00"
         Mask            =   "##.##"
         PromptChar      =   "0"
      End
      Begin VB.TextBox TextCiudad 
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
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2730
         Width           =   1380
      End
      Begin VB.TextBox TextEmpresa 
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
         MaxLength       =   40
         TabIndex        =   11
         Top             =   2730
         Width           =   4110
      End
      Begin VB.TextBox TextCliente 
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
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2205
         Width           =   5160
      End
      Begin VB.CommandButton Command4 
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
         Height          =   960
         Left            =   7770
         Picture         =   "Clientes.frx":0718
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1995
         Width           =   1065
      End
      Begin VB.CheckBox CheqContrato 
         Caption         =   "Contrato No."
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
         TabIndex        =   38
         Top             =   5040
         Width           =   1485
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   39
         Top             =   5250
         Visible         =   0   'False
         Width           =   8835
         Begin VB.TextBox TextCant 
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
            Left            =   7980
            MaxLength       =   5
            TabIndex        =   52
            Top             =   525
            Width           =   750
         End
         Begin VB.TextBox TextArea 
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
            Left            =   7245
            MaxLength       =   5
            TabIndex        =   50
            Top             =   525
            Width           =   750
         End
         Begin VB.TextBox TextSector 
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
            Left            =   5355
            MaxLength       =   20
            TabIndex        =   48
            Top             =   525
            Width           =   1905
         End
         Begin VB.TextBox TextContrato 
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
            Left            =   4095
            MaxLength       =   8
            TabIndex        =   46
            Top             =   525
            Width           =   1275
         End
         Begin VB.OptionButton OpcN 
            Caption         =   "Nueva"
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
            TabIndex        =   40
            Top             =   525
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton OpcR 
            Caption         =   "Renovación"
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
            TabIndex        =   41
            Top             =   210
            Width           =   1380
         End
         Begin MSMask.MaskEdBox MBoxHasta 
            Height          =   330
            Left            =   2835
            TabIndex        =   44
            Top             =   525
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
         Begin MSMask.MaskEdBox MBoxDesde 
            Height          =   330
            Left            =   1575
            TabIndex        =   43
            Top             =   525
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
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cant."
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
            TabIndex        =   51
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Area"
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
            TabIndex        =   49
            Top             =   210
            Width           =   750
         End
         Begin VB.Label Label15 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Sector"
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
            TabIndex        =   47
            Top             =   210
            Width           =   1905
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Contrato No."
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
            TabIndex        =   45
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Periodo"
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
            TabIndex        =   42
            Top             =   210
            Width           =   2535
         End
      End
      Begin VB.CheckBox CheqMult 
         Caption         =   "Factura Múltiple"
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
         Left            =   6615
         TabIndex        =   16
         Top             =   2520
         Width           =   1170
      End
      Begin MSMask.MaskEdBox MBoxRUCE 
         Height          =   330
         Left            =   -69855
         TabIndex        =   64
         Top             =   2205
         Width           =   2220
         _ExtentX        =   3916
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
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PAIS"
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
         Left            =   -69855
         TabIndex        =   78
         Top             =   2625
         Width           =   2430
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DIAS DE HOSP."
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
         Left            =   -74895
         TabIndex        =   76
         Top             =   1995
         Width           =   1590
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CANT. HABIT."
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
         Left            =   -73320
         TabIndex        =   75
         Top             =   1995
         Width           =   1800
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " C.I. / R.U.C."
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
         Left            =   -69855
         TabIndex        =   72
         Top             =   1995
         Width           =   2220
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. PLACAS VEHICULO / ID. CAR"
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
         Left            =   -74895
         TabIndex        =   71
         Top             =   3255
         Width           =   3480
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TARJETA"
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
         Left            =   -74895
         TabIndex        =   70
         Top             =   2625
         Width           =   2010
      End
      Begin VB.Label Label31 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TARJETA DE CREDITO No."
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
         Left            =   -72900
         TabIndex        =   69
         Top             =   2625
         Width           =   2745
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NACIONALIDAD:"
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
         Left            =   -69855
         TabIndex        =   68
         Top             =   3255
         Width           =   2430
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EST. CIVIL:"
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
         Left            =   -74895
         TabIndex        =   59
         Top             =   1995
         Width           =   2640
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PAIS"
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
         Left            =   -72270
         TabIndex        =   58
         Top             =   1995
         Width           =   2325
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GRUPO:"
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
         Left            =   5460
         TabIndex        =   36
         Top             =   4410
         Width           =   1065
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROFESION"
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
         TabIndex        =   34
         Top             =   4410
         Width           =   4110
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA NAC."
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
         TabIndex        =   32
         Top             =   4410
         Width           =   1275
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección:"
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
         TabIndex        =   29
         Top             =   3570
         Width           =   5160
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telefono(s):"
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
         TabIndex        =   26
         Top             =   3570
         Width           =   1275
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FAX:"
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
         Left            =   5250
         TabIndex        =   21
         Top             =   3045
         Width           =   1275
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " R.U.C / C.I."
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
         Left            =   3465
         TabIndex        =   19
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " E-MAIL"
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
         TabIndex        =   17
         Top             =   3045
         Width           =   3375
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comision"
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
         Left            =   5565
         TabIndex        =   14
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label LabelCodigo 
         BackColor       =   &H80000009&
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
         Left            =   5250
         TabIndex        =   9
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad:"
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
         Left            =   4200
         TabIndex        =   12
         Top             =   2520
         Width           =   1380
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMPRESA:"
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
         TabIndex        =   10
         Top             =   2520
         Width           =   4110
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO:"
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
         Left            =   5250
         TabIndex        =   8
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CLIENTE:"
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
         TabIndex        =   6
         Top             =   1995
         Width           =   5160
      End
   End
   Begin MSAdodcLib.Adodc AdoGrupoNo 
      Height          =   330
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "GrupoNo"
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
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin MSAdodcLib.Adodc AdoDBCliente 
      Height          =   330
      Left            =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "DBCliente"
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
   Begin MSAdodcLib.Adodc AdoContratos 
      Height          =   330
      Left            =   0
      Top             =   1800
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
      Caption         =   "Contratos"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":0B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":0C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":0D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":0E90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":13A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":18B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1DC6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqContrato_Click()
 If CheqContrato.Value = 1 Then
    Frame2.Visible = True
 Else
    Frame2.Visible = False
 End If
End Sub

Private Sub Command1_Click()
   Unload FClientes
End Sub

Private Sub Command4_Click()
   FechaValida MBoxFecha_N
   TextoValido TextEmpresa, , False
   TextoValido TextCiudad, , False
   TextoValido TextEmail, , False
   TextoValido TextProfesion, , False
   TextoValido TextDireccion, , False
   TextoValido TextDireccion1, , False
   TextoValido TextSector, , True
   TextoValido TextArea, , True
   TextoValido TextGrupo
   TextoValido TextCant, True
   If OpcCliente.Value Then
      Codigo = "C"
   ElseIf OpcEjecutivo.Value Then
      Codigo = "E"
   Else
      Codigo = "R"
   End If
   Grupo_No = DCGrupoNo.Text
   GrabarCliente Codigo
   sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
   sSQL = sSQL & "FROM Clientes "
   sSQL = sSQL & "WHERE T = '" & Codigo & "' "
   If Codigo = "C" Then sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
   sSQL = sSQL & "ORDER BY Cliente "
   SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
   LlenarCliente DLCliente.Text, Codigo
End Sub

Private Sub DCGrupoNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupoNo_LostFocus()
 Grupo_No = DCGrupoNo.Text
End Sub

Private Sub DLCliente_DblClick()
  SiguienteControl
End Sub

Private Sub DLCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCliente_LostFocus()
  If OpcCliente.Value Then
     Codigo = "C"
  ElseIf OpcEjecutivo.Value Then
     Codigo = "E"
  Else
     Codigo = "R"
  End If
  If AdoDBCliente.Recordset.RecordCount > 0 Then
     LlenarCliente DLCliente.Text, Codigo
  End If
End Sub

Private Sub Form_Activate()
  Grabar = True
  sSQL = "SELECT Grupo "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "GROUP BY Grupo "
  SelectDBCombo DCGrupoNo, AdoGrupoNo, sSQL, "Grupo"
  Grupo_No = DCGrupoNo.Text
  sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE T = 'C' "
  sSQL = sSQL & "AND Grupo = '" & NumEmpresa & "' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
  LlenarCliente DLCliente.Text, Codigo
  Label12.Caption = "Descuento"
  RatonNormal
End Sub

Private Sub Form_Load()
CentrarForm FClientes
ConectarAdodc AdoGrupoNo
ConectarAdodc AdoCliente
ConectarAdodc AdoContratos
ConectarAdodc AdoDBCliente
End Sub

Private Sub MBoxDesde_LostFocus()
  FechaValida MBoxDesde
End Sub

Private Sub MBoxFecha_N_LostFocus()
  FechaValida MBoxFecha_N
End Sub

Private Sub MBoxHasta_LostFocus()
  FechaValida MBoxHasta
End Sub

Private Sub OpcCliente_Click()
  SiguienteControl
End Sub

Private Sub OpcCliente_DblClick()
  SiguienteControl
End Sub

Private Sub OpcCliente_LostFocus()
  Grupo_No = DCGrupoNo.Text
  sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE T = 'C' "
  sSQL = sSQL & "AND Grupo = '" & NumEmpresa & "' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
  LlenarCliente DLCliente.Text, "C"
  Label12.Caption = "Descuento"
End Sub

Private Sub OpcEjecutivo_Click()
  sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE E = 'E' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
  LlenarCliente DLCliente.Text, "E"
  Label12.Caption = "Comisión"
End Sub

Private Sub OpcResp_Click()
  sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE E = 'R' "
  sSQL = sSQL & "ORDER BY Cliente "
  SelectDBList DLCliente, DataDBCliente, sSQL, "NomClientes"
  LlenarCliente DLCliente.Text, "R"
  Label12.Caption = "Comisión"
End Sub

Private Sub TextPlacas_LostFocus()
  TextoValido TextPlacas, , True
End Sub

Private Sub TextTarj_LostFocus()
  TextoValido TextTarj, False, True
End Sub

Private Sub TextTarjCred_LostFocus()
  TextoValido TextTarjCred, False, True
End Sub

Private Sub TextArea_LostFocus()
  TextoValido TextArea, , True
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
End Sub

Private Sub TextCant_LostFocus()
  TextoValido TextCant, True
End Sub

Private Sub TextCiudad_GotFocus()
 MarcarTexto TextCiudad
End Sub

Private Sub TextCiudad_LostFocus()
 TextoValido TextCiudad, , False
End Sub

Private Sub TextContrato_LostFocus()
  TextoValido TextContrato, , True
End Sub

Private Sub TextDireccion_GotFocus()
  MarcarTexto TextDireccion
End Sub

Private Sub TextDireccion_LostFocus()
 TextoValido TextDireccion, , False
End Sub

Private Sub TextDireccion1_GotFocus()
  MarcarTexto TextDireccion1
End Sub

Private Sub TextDireccion1_LostFocus()
 TextoValido TextDireccion1, , False
End Sub

Private Sub TextEmail_LostFocus()
  TextoValido TextEmail, , False
End Sub

Private Sub TextEmpresa_LostFocus()
  TextoValido TextEmpresa, , False
End Sub

Public Sub LlenarCliente(Clientes As String, OpcCli As String)
   'If Nuevo = False Then
   Codigo1 = SinEspaciosDer(Clientes)
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE Codigo = '" & Codigo1 & "' "
   sSQL = sSQL & "AND T = '" & OpcCli & "' "
   SelectData AdoCliente, sSQL
   With AdoCliente.Recordset
   If .RecordCount > 0 Then
       TextCliente.Text = .Fields("Cliente")
       TextEmpresa.Text = .Fields("Empresa")
       LabelCodigo.Caption = .Fields("Codigo")
       TextDireccion.Text = .Fields("Direccion")
       TextProfesion.Text = .Fields("Profesion")
       MBoxFecha_N.Text = Format(.Fields("Fecha_N"), FormatoFechas)
       TextDireccion1.Text = .Fields("Direccion1")
       MBoxTelef1.Text = FormatoCodigoTelef(.Fields("Telefono"))
       MBoxTelef2.Text = FormatoCodigoTelef(.Fields("Celular"))
       MBoxFAX.Text = FormatoCodigoTelef(.Fields("FAX"))
       TextCiudad.Text = .Fields("Ciudad")
       TextEmail.Text = .Fields("Email")
       MBoxRUC.Text = FormatoCodigoRUC_CI(.Fields("RUC_CI"), True)
       If .Fields("FactM") Then
           CheqMult.Value = 1
       Else
           CheqMult.Value = 0
       End If
       If .Fields("Sexo") = "M" Then
           OpcH.Value = True
       Else
           OpcM.Value = True
       End If
       TextGrupo.Text = .Fields("Grupo")
       MBoxComision.Text = Format(.Fields("Porc_C"), "00.00")
       sSQL = "SELECT * FROM Contratos_Suscrip "
       sSQL = sSQL & "WHERE Codigo_C = '" & LabelCodigo.Caption & "' "
       sSQL = sSQL & "AND T <> 'A' "
       SelectData AdoContratos, sSQL
       With AdoContratos.Recordset
        If .RecordCount > 0 Then
            CheqContrato.Value = 1
            TextContrato.Text = .Fields("Contrato_No")
            MBoxDesde.Text = Format(.Fields("Desde"), FormatoFechas)
            MBoxHasta.Text = Format(.Fields("Hasta"), FormatoFechas)
            TextSector.Text = .Fields("Sector")
            TextArea.Text = .Fields("Area")
            TextCant.Text = .Fields("Contador")
            If .Fields("Nuevo") Then
               OpcN.Value = True
            Else
               OpcR.Value = True
            End If
            Select Case .Fields("T")
              Case "N": Frame2.Caption = "   Estado del Contrato: Pendiente "
              Case "A": Frame2.Caption = "   Estado del Contrato: Anulado "
              Case "R": Frame2.Caption = "   Estado del Contrato: Renovación "
              Case "S": Frame2.Caption = "   Estado del Contrato: Suspendido "
              Case "C": Frame2.Caption = "   Estado del Contrato: Cancelado "
            End Select
            Frame2.Visible = True
        Else
            CheqContrato.Value = 0
        End If
       End With
   End If
   End With
   'End If
End Sub

Private Sub TextGrupo_GotFocus()
  MarcarTexto TextGrupo
End Sub

Private Sub TextGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextGrupo_LostFocus()
  TextoValido TextGrupo
End Sub

Private Sub TextProfesion_LostFocus()
  TextoValido TextProfesion
End Sub

Private Sub TextSector_LostFocus()
  TextoValido TextSector, , True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  Nuevo = False
  With AdoDBCliente.Recordset
  Select Case Button.key
    Case "Grabar"
         If OpcCliente.Value Then
            Codigo = "C"
         ElseIf OpcEjecutivo.Value Then
            Codigo = "E"
         Else
            Codigo = "R"
         End If
         GrabarCliente Codigo
         Grupo_No = DCGrupoNo.Text
         sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
         sSQL = sSQL & "FROM Clientes "
         sSQL = sSQL & "WHERE T = '" & Codigo & "' "
         If Codigo = "C" Then sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
         sSQL = sSQL & "ORDER BY Cliente "
         SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
         Nuevo = False
    Case "Eliminar"
         If OpcCliente.Value Then
            Codigo = "C"
         ElseIf OpcEjecutivo.Value Then
            Codigo = "E"
         Else
            Codigo = "R"
         End If
         Codigos = LabelCodigo.Caption
         Mensajes = "Esta seguro que desea eliminar," & Chr$(13)
         Mensajes = Mensajes & "El Cliente: " & DLCliente.Text & "          (" & Codigos & ")."
         Titulo = "Eliminacion de Clientes"
         TipoDeCaja = 4 + 32: J = MsgBox(Mensajes, TipoDeCaja, Titulo)
         If J = 6 Then
            sSQL = "DELETE * FROM Clientes "
            sSQL = sSQL & "WHERE Codigo = '" & Codigos & "';"
            ConectarAdoExecute sSQL
         End If
         Grupo_No = Val(DBCGrupoNo.Text)
         sSQL = "SELECT Cliente & Space(90-Len(Cliente)) & Codigo As NomClientes "
         sSQL = sSQL & "FROM Clientes "
         sSQL = sSQL & "WHERE E = '" & Codigo & "' "
         If Codigo = "C" Then sSQL = sSQL & "AND Grupo = '" & Grupo_No & "' "
         sSQL = sSQL & "ORDER BY Cliente "
         SelectDBList DLCliente, AdoDBCliente, sSQL, "NomClientes"
        .MoveFirst
    Case "Nuevo"
         Nuevo = True
         LabelCodigo.Caption = "Codigo"
         TextEmpresa.Text = ""
         TextProfesion.Text = ""
         TextDireccion.Text = ""
         MBoxTelef1.Text = "00-000-000"
         MBoxTelef2.Text = "00-000-000"
         MBoxFAX.Text = "00-000-000"
         MBoxRUC.Text = "000000000-0-000"
         TextCiudad.Text = ""
         TextCliente.Text = ""
         TextCliente.SetFocus
    Case "Primero"
        .MoveFirst
    Case "Anterior"
        .MovePrevious
         If .BOF Then .MoveFirst
    Case "Siguiente"
        .MoveNext
         If .EOF Then .MoveLast
    Case "Ultimo"
        .MoveLast
  End Select
  If AdoDBCliente.Recordset.RecordCount > 0 Then
     '.MoveFirst
      If OpcCliente.Value Then
         Codigo = "C"
      ElseIf OpcEjecutivo.Value Then
         Codigo = "E"
      Else
         Codigo = "R"
      End If
      DLCliente.Text = AdoDBCliente.Recordset.Fields("NomClientes")
      If Not Nuevo Then LlenarCliente DLCliente.Text, Codigo
  End If
  End With
End Sub

Public Sub GrabarCliente(OpcCli As String)
Mensajes = "Esta Seguro que desea grabar el Cliente: " & Chr(13) & "["
Mensajes = Mensajes & TextCliente.Text & "]"
Titulo = "Pregunta de Grabación"
If BoxMensaje = vbYes Then
   sSQL = "SELECT * FROM Clientes "
   sSQL = sSQL & "WHERE Cliente = '" & TextCliente.Text & "' "
   sSQL = sSQL & "AND T = '" & OpcCli & "' "
   SelectData AdoCliente, sSQL
   With AdoCliente.Recordset
     If .RecordCount > 0 Then
        '.Edit
         LabelCodigo.Caption = .Fields("Codigo")
         If .Fields("Grupo") <> TextGrupo.Text Then
             Mensajes = "Quiere cambiar de Grupo"
             Titulo = "Pregunta de Cambio de Grupo"
             If BoxMensaje = vbNo Then
               .AddNew
                Numero = ReadSetDataNum("Clientes", True, True)
                LabelCodigo.Caption = FormatoCodigo(TextCliente.Text, Numero)
             End If
         End If
     Else
        .AddNew
         Numero = ReadSetDataNum("Clientes", True, True)
         LabelCodigo.Caption = FormatoCodigo(TextCliente.Text, Numero)
     End If
       .Fields("Codigo") = LabelCodigo.Caption
       .Fields("Cliente") = TextCliente.Text
       .Fields("Actividad") = TextEmpresa.Text
       .Fields("Profesion") = TextProfesion.Text
       .Fields("Fecha_N") = MBoxFecha_N.Text
       .Fields("Direccion") = TextDireccion.Text
       .Fields("DireccionT") = TextDireccion1.Text
       .Fields("Telefono") = MBoxTelef1.Text
       .Fields("Celular") = MBoxTelef2.Text
       .Fields("FAX") = MBoxFAX.Text
       .Fields("Ciudad") = TextCiudad.Text
       .Fields("Email") = TextEmail.Text
       .Fields("CI_RUC") = MBoxRUC.Text
       .Fields("Porc_C") = CSng(MBoxComision.Text)
       .Fields("Grupo") = TextGrupo.Text
       .Fields("FactM") = False
       .Fields("Sexo") = "F"
        If CheqMult.Value = 1 Then .Fields("FactM") = True
        If OpcH.Value Then .Fields("Sexo") = "M"
       .Fields("T") = OpcCli
       .Update
        Nuevo = False
   End With
   If CheqContrato.Value = 1 Then
     FechaValida MBoxDesde, False
     FechaValida MBoxHasta, False
     TextoValido TextContrato, False, True
     If TextContrato.Text <> Ninguno Then
        sSQL = "SELECT * FROM Contratos_Suscrip "
        sSQL = sSQL & "WHERE Contrato_No = '" & TextContrato.Text & "' "
        SelectData AdoContratos, sSQL
        With AdoContratos.Recordset
         If .RecordCount > 0 Then
            '.Edit
         Else
            .AddNew
            .Fields("Contrato_No") = Mid(TextContrato.Text, 1, 8)
         End If
             If OpcN.Value Then
               .Fields("Nuevo") = True
               .Fields("Fecha_I") = MBoxDesde.Text
               .Fields("T") = Normal
             Else
               .Fields("Nuevo") = False
               .Fields("T") = Renovacion
             End If
            .Fields("Codigo_C") = LabelCodigo.Caption
            .Fields("Desde") = MBoxDesde.Text
            .Fields("Hasta") = MBoxHasta.Text
            .Fields("Sector") = TextSector.Text
            .Fields("Area") = TextArea.Text
            .Fields("Contador") = Val(TextCant.Text)
            .Fields("Factura_No") = 0
            .Update
        End With
     End If
   End If
End If
End Sub

