VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{64AED23E-31A2-4023-8C7D-E628B15843D8}#1.0#0"; "Code39X.ocx"
Begin VB.Form IngActivos 
   Caption         =   "Ingreso/Modificacion de Productos de Inventario"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Height          =   750
      Left            =   10815
      Picture         =   "IngActiv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1680
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Imprimir"
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
      Left            =   10815
      Picture         =   "IngActiv.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   840
      Width           =   960
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   105
      TabIndex        =   0
      Top             =   3150
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "IngActiv.frx":0D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblTotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DCResp"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MBoxCta_Pat"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "MBFecha"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Option1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Option2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Picture1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "MBoxCodigo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtSubCta"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MBoxCta_Inv"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtUnidad"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtBarra"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtFactura"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtSubTotal"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtIVA"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtDetalle"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtCantidad"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "DCProv"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CheqConta"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Detalle del Producto"
      TabPicture(1)   =   "IngActiv.frx":1026
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.CheckBox CheqConta 
         Caption         =   "Contabilizar"
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
         Left            =   10185
         TabIndex        =   5
         Top             =   735
         Width           =   1380
      End
      Begin MSDataListLib.DataCombo DCProv 
         Bindings        =   "IngActiv.frx":1900
         DataSource      =   "AdoProv"
         Height          =   315
         Left            =   6090
         TabIndex        =   11
         Top             =   1050
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
      Begin VB.TextBox TxtCantidad 
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
         Left            =   9660
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox TxtDetalle 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1695
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Text            =   "IngActiv.frx":1916
         Top             =   2415
         Width           =   7890
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
         Height          =   330
         Left            =   9660
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   2310
         Width           =   1905
      End
      Begin VB.TextBox TxtSubTotal 
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
         Left            =   9660
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1995
         Width           =   1905
      End
      Begin VB.TextBox TxtFactura 
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
         Left            =   9660
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1365
         Width           =   1905
      End
      Begin VB.TextBox TxtBarra 
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
         Left            =   4830
         MaxLength       =   25
         TabIndex        =   21
         ToolTipText     =   "<Alt+F2> Codigo Automático"
         Top             =   1995
         Width           =   3165
      End
      Begin VB.TextBox TxtUnidad 
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
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1995
         Width           =   2640
      End
      Begin MSMask.MaskEdBox MBoxCta_Inv 
         Height          =   330
         Left            =   2835
         TabIndex        =   13
         Top             =   1365
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
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
      Begin VB.TextBox TxtSubCta 
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
         MaxLength       =   60
         TabIndex        =   4
         Top             =   735
         Width           =   4950
      End
      Begin MSMask.MaskEdBox MBoxCodigo 
         Height          =   330
         Left            =   1155
         TabIndex        =   2
         Top             =   735
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   "CC.CC.CCC.CCCCC"
         Mask            =   "CC.CC.CCC.CCCCC"
         PromptChar      =   " "
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   8085
         ScaleHeight     =   855
         ScaleWidth      =   3480
         TabIndex        =   36
         Top             =   3045
         Width           =   3480
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Detalle"
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
         TabIndex        =   7
         Top             =   1365
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agrupación"
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
         Top             =   1050
         Value           =   -1  'True
         Width           =   1380
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   3465
         TabIndex        =   9
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
      Begin MSMask.MaskEdBox MBoxCta_Pat 
         Height          =   330
         Left            =   6090
         TabIndex        =   15
         Top             =   1365
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
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
      Begin MSDataListLib.DataCombo DCResp 
         Bindings        =   "IngActiv.frx":191A
         DataSource      =   "AdoResp"
         Height          =   315
         Left            =   1470
         TabIndex        =   17
         Top             =   1680
         Width           =   6525
         _ExtentX        =   11509
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
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVEEDOR:"
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
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. &PATRIM."
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
         TabIndex        =   14
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
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
         Left            =   8085
         TabIndex        =   24
         Top             =   1680
         Width           =   1590
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVEEDOR:"
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
         TabIndex        =   10
         Top             =   1050
         Width           =   1380
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
         Height          =   330
         Left            =   9660
         TabIndex        =   31
         Top             =   2625
         Width           =   1905
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Producto"
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
         Left            =   8085
         TabIndex        =   30
         Top             =   2625
         Width           =   1590
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total I.V.A."
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
         Left            =   8085
         TabIndex        =   28
         Top             =   2310
         Width           =   1590
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sub Total"
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
         Left            =   8085
         TabIndex        =   26
         Top             =   1995
         Width           =   1590
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &FACTURA No."
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
         Left            =   8085
         TabIndex        =   22
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA DE COMPRA:"
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
         TabIndex        =   8
         Top             =   1050
         Width           =   2010
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " B&ARRA:"
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
         TabIndex        =   20
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &UBICACION:"
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
         TabIndex        =   18
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CTA. &ACTIVO"
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
         TabIndex        =   12
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DETALLE DEL ACTIVO"
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
         TabIndex        =   3
         Top             =   735
         Width           =   2220
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Codigo:"
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
         Top             =   735
         Width           =   1065
      End
   End
   Begin ComctlLib.TreeView TVCatalogo 
      Height          =   3000
      Left            =   105
      TabIndex        =   33
      ToolTipText     =   "Un click en el dibujo de la Cta. y presionar la tecla <DEL> Borra la Cta."
      Top             =   105
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   5292
      _Version        =   327682
      Indentation     =   794
      Sorted          =   -1  'True
      Style           =   5
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
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
   Begin Code39X.Code39Clt Code39Clt1 
      Left            =   6090
      Top             =   420
      _ExtentX        =   1905
      _ExtentY        =   1085
   End
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   420
      Top             =   1260
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "Inv"
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
      Height          =   750
      Left            =   10815
      Picture         =   "IngActiv.frx":1930
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   945
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
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   420
      Top             =   1575
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoCodInv 
      Height          =   330
      Left            =   420
      Top             =   1890
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
      Caption         =   "CodInv"
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
   Begin MSAdodcLib.Adodc AdoProv 
      Height          =   330
      Left            =   420
      Top             =   2205
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
      Caption         =   "Prov"
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
   Begin MSAdodcLib.Adodc AdoResp 
      Height          =   330
      Left            =   420
      Top             =   630
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
      Caption         =   "Resp"
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
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngActiv.frx":1D72
            Key             =   "Uno"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngActiv.frx":208C
            Key             =   "Dos"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "IngActiv.frx":23A6
            Key             =   "Tres"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "IngActivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cta_Ini As String
Dim Cta_Fin As String
Dim nodX As Node

Public Sub AddNewCtaInv(TipoTC As String)
Dim InicioCodigo As Byte
 'MsgBox Codigo
  InicioCodigo = 1
  For IE = 1 To Len(MascaraCodigoA)
    If Mid(MascaraCodigoA, IE, 1) = "." Then
       InicioCodigo = IE
       IE = Len(MascaraCodigoA) + 1
    End If
  Next IE
  If Len(Codigo) = InicioCodigo Then
     Set nodX = TVCatalogo.Nodes.Add(, , Codigo, Cuenta)
     nodX.Image = ImageList1.ListImages(1).key
     nodX.SelectedImage = ImageList1.ListImages(1).key
  Else
     Set nodX = TVCatalogo.Nodes.Add(Cta_Sup, tvwChild, Codigo, Cuenta)
     Select Case TipoTC
       Case "I": IE = 1
       Case "P": IE = 2
       Case Else: IE = 3
     End Select
     nodX.Image = ImageList1.ListImages(IE).key
     nodX.SelectedImage = ImageList1.ListImages(IE).key
     TVCatalogo.Tag = Codigo
  End If
End Sub

Public Sub ColocarCodigoBarra(CodigoDeBarra As String)
Dim PosSup, PosIzq As Single
  Code39Clt1.AlturaBarra = 45
  Code39Clt1.TamBarra = 1
  Code39Clt1.ColorCodigo = "N"
  Code39Clt1.ValorCodigo = CodigoDeBarra
  Code39Clt1.RealizarCodigo
  
  Picture1 = Clipboard.GetData
  Picture1.FontBold = True
  PosIzq = ((Picture1.width - Picture1.TextWidth(Code39Clt1.ValorCodigo)) / 2)
  PosSup = Picture1.Height - 360
  If PosIzq < 0 Then PosIzq = 0.1
  If PosSup < 0 Then PosSup = 0.1
  Picture1.Line (150, PosSup + 140)-(Picture1.width - 150, Picture1.Height), QBColor(Blanco_Brillante), BF
  Picture1.CurrentX = PosIzq
  Picture1.CurrentY = PosSup + 150
  Picture1.Print Code39Clt1.ValorCodigo
  PosSup = Picture1.Height - 160
''  Cadena = "USD$ " & Format(PVP, "#,##0.0000")
''  PosIzq = ((Picture1.Width - Picture1.TextWidth(Cadena)) / 2)
''  Picture1.CurrentX = PosIzq
''  Picture1.CurrentY = PosSup
''  Picture1.Print Cadena
  Picture1.FontBold = False
  If TxtBarra.Text = Ninguno Then TxtBarra.Text = Code39Clt1.ValorCodigo
End Sub

Private Sub Command1_Click()
  GrabarInv
End Sub

Private Sub Command2_Click()
  Unload IngActivos
End Sub

Private Sub Command6_Click()
  Imprimir_Codigos_Estanteria MBoxCodigo, Code39Clt1, Picture1
End Sub

Private Sub DCProv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProv_LostFocus()
  If AdoProv.Recordset.RecordCount > 0 Then
     AdoProv.Recordset.MoveFirst
     AdoProv.Recordset.Find ("Cliente = '" & DCProv.Text & "' ")
     If Not AdoProv.Recordset.EOF Then
        CodigoCliente = AdoProv.Recordset.Fields("Codigo")
     Else
        MsgBox "PROVEEDOR NO ASIGNADO"
        CodigoCliente = Ninguno
     End If
  End If
End Sub

Private Sub DCResp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCResp_LostFocus()
  If AdoResp.Recordset.RecordCount > 0 Then
     AdoResp.Recordset.MoveFirst
     AdoResp.Recordset.Find ("Cliente = '" & DCResp.Text & "' ")
     If Not AdoResp.Recordset.EOF Then
        CodigoBenef = AdoResp.Recordset.Fields("Codigo")
     Else
        MsgBox "RESPONSABLE NO ASIGNADO"
        CodigoBenef = Ninguno
     End If
  End If
End Sub

Private Sub Form_Activate()
  FormatoMaskCta MBoxCta_Inv
  FormatoMaskCta MBoxCta_Pat
  FormatoMaskCodA MBoxCodigo
  Si_No = False
  
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  SelectDBCombo DCProv, AdoProv, sSQL, "Cliente"
  SelectDBCombo DCResp, AdoResp, sSQL, "Cliente"
  
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Producto = UCASE(Producto) " _
       & "WHERE TC = 'I' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Producto <> UCASE(Producto) "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT Item,Codigo_Inv,Codigo_Barra,Periodo " _
       & "FROM Catalogo_Productos " _
       & "WHERE TC = 'P' " _
       & "AND TDP = 'ACT' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoTInv, sSQL
  If AdoTInv.Recordset.RecordCount > 0 Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TDP = 'ACT' " _
          & "ORDER BY Codigo_Inv "
     SelectAdodc AdoInv, sSQL
     If AdoInv.Recordset.RecordCount > 0 Then
        ReDim CodigoCtas(AdoInv.Recordset.RecordCount + 1) As String
        For I = 0 To AdoInv.Recordset.RecordCount
            CodigoCtas(I) = ""
        Next I
     End If
     Contador = 0
     Do While Not AdoTInv.Recordset.EOF
        Codigo = AdoTInv.Recordset.Fields("Codigo_Inv")
        Codigo1 = AdoTInv.Recordset.Fields("Codigo_Barra")
        If Codigo1 = Ninguno Then
           Codigo1 = AdoTInv.Recordset.Fields("Codigo_Inv")
           Codigo1 = Replace(Codigo1, ".", "")
           AdoTInv.Recordset.Fields("Codigo_Barra") = Codigo1
           AdoTInv.Recordset.Update
        End If
        Cta_Sup = CodigoCuentaSup(Codigo)
        With AdoInv.Recordset
         If .RecordCount > 0 Then
             Do While (Cta_Sup <> "0")
               .MoveFirst
               .Find ("Codigo_Inv Like '" & Cta_Sup & "' ")
                If Not .EOF And Cta_Sup <> "0" Then
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                Else
                   Si_No = True: Evaluar = True
                   For I = 0 To AdoInv.Recordset.RecordCount
                       If CodigoCtas(I) = Cta_Sup Then Evaluar = False
                   Next I
                   If Evaluar Then
                      SetAdoAddNew "Catalogo_Productos"
                      SetAdoFields "TDP", "ACT"
                      SetAdoFields "Item", NumEmpresa
                      SetAdoFields "Codigo_Inv", Cta_Sup
                      SetAdoFields "Producto", "PRODUCTO SIN NOMBRE"
                      SetAdoFields "Periodo", Periodo_Contable
                      SetAdoFields "TC", "I"
                      SetAdoUpdate
                      CodigoCtas(Contador) = Cta_Sup
                      Contador = Contador + 1
                   End If
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                End If
             Loop
         End If
        End With
        AdoTInv.Recordset.MoveNext
     Loop
  End If
  RatonNormal
  If Si_No Then
     Cadena = vbCrLf
     For I = 0 To Contador
         Cadena = Cadena & CodigoCtas(I) & vbCrLf
     Next I
     MsgBox "Los siguientes Codigos no se han creado: " & vbCrLf _
            & Cadena & "ADVERTENCIA: REVIZAR."
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TDP = 'ACT' " _
       & "ORDER BY Codigo_Activo "
  SelectAdodc AdoInv, sSQL
  RatonReloj
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          If Len(.Fields("Codigo_Activo")) = 2 Then
             Codigo = "C" & .Fields("Codigo_Activo")
             Cta_Sup = "C" & .Fields("Codigo_Activo")
             Cuenta = .Fields("Codigo_Activo") & " - " & .Fields("Producto")
             AddNewCtaInv .Fields("TC")
          Else
             Codigo = "C" & .Fields("Codigo_Activo")
             Cta_Sup = "C" & CodigoCuentaSup(.Fields("Codigo_Activo"))
             Cuenta = .Fields("Codigo_Activo") & " - " & .Fields("Producto")
             AddNewCtaInv .Fields("TC")
          End If
         .MoveNext
       Loop
   End If
  End With
  MBoxCodigo.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoInv
  ConectarAdodc AdoTInv
  ConectarAdodc AdoProv
  ConectarAdodc AdoResp
  ConectarAdodc AdoCodInv
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCodigo_GotFocus()
  MarcarTexto MBoxCodigo
End Sub

Private Sub MBoxCodigo_LostFocus()
   MBoxCodigo.Text = UCase(MBoxCodigo.Text)
End Sub

Private Sub MBoxCta_Inv_GotFocus()
  MarcarTexto MBoxCta_Inv
End Sub

Private Sub MBoxCta_Inv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_Pat_GotFocus()
  MarcarTexto MBoxCta_Pat
End Sub

Private Sub MBoxCta_Pat_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBarra_GotFocus()
  MarcarTexto TxtBarra
End Sub

Private Sub TxtBarra_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If AltDown And KeyCode = vbKeyF2 Then
     RatonReloj
     sSQL = "SELECT MAX(Codigo_Barra) As Maximo " _
          & "FROM Catalogo_Productos " _
          & "WHERE TC = 'P' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     SelectData AdoCodInv, sSQL
     With AdoCodInv.Recordset
      If .RecordCount > 0 Then
          TxtBarra = Format(Val(.Fields("Maximo")) + 1, "0000")
      Else
          TxtBarra = "0001"
      End If
     End With
     RatonNormal
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtBarra_LostFocus()
  TextoValido TxtBarra
End Sub

Private Sub TxtCantidad_GotFocus()
  MarcarTexto TxtCantidad
End Sub

Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFactura_GotFocus()
  MarcarTexto TxtFactura
End Sub

Private Sub TxtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIVA_GotFocus()
  MarcarTexto TxtIVA
End Sub

Private Sub TxtIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIVA_LostFocus()
  LblTotal.Caption = Format(CCur(Val(TxtSubTotal)) + CCur(Val(TxtIVA)), "#,##0.00")
End Sub

Private Sub TxtSubCta_GotFocus()
  MarcarTexto TxtSubCta
End Sub

Private Sub TxtSubCta_LostFocus()
  TextoValido TxtSubCta
End Sub

Public Sub LlenarInv()
   FormatoMaskCta MBoxCta_Inv
   FormatoMaskCta MBoxCta_Pat
   FormatoMaskCodA MBoxCodigo
   With AdoInv.Recordset
    If .RecordCount > 0 Then
        Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
       .MoveFirst
        TextoBusqueda = "Codigo_Activo Like '" & Codigo & "' "
       .Find (TextoBusqueda)
        If Not .EOF Then
           If .Fields("T") = "C" Then CheqConta.value = 1 Else CheqConta.value = 0
           TxtSubCta.Text = .Fields("Producto")
           TxtUnidad.Text = .Fields("Ubicacion")
           TxtBarra.Text = .Fields("Codigo_Barra")
           If TxtBarra.Text = Ninguno Then
              TxtBarra.Text = CodigosSinPuntos(.Fields("Codigo_Activo"))
           End If
           ColocarCodigoBarra TxtBarra.Text
          'MsgBox FormatoCodigoActivo(.Fields("Codigo_Activo"))
           MBoxCodigo.Text = FormatoCodigoActivo(.Fields("Codigo_Activo"))
           MBoxCta_Inv.Text = FormatoCodigoCta(.Fields("Cta_Activo"))
           MBoxCta_Pat.Text = FormatoCodigoCta(.Fields("Cta_Patrimonio"))
           If .Fields("TC") = "P" Then
               Option1.value = False
               Option2.value = True
           Else
               Option1.value = True
               Option2.value = False
               TxtBarra.Text = "0000000000"
           End If
           TxtDetalle.Text = .Fields("Detalle")
           MBFecha.Text = .Fields("Fecha_Compra")
           TxtFactura.Text = .Fields("Factura_No")
           TxtCantidad = Format(.Fields("Cantidad"), "#,##0.00")
           TxtSubTotal.Text = Format(.Fields("Sub_Total"), "#,##0.00")
           TxtIVA.Text = Format(.Fields("Total_IVA"), "#,##0.00")
           LblTotal.Caption = Format(.Fields("Total_Compra"), "#,##0.00")
           If AdoProv.Recordset.RecordCount > 0 Then
              AdoProv.Recordset.MoveFirst
              AdoProv.Recordset.Find ("Codigo = '" & .Fields("Codigo_Prov") & "' ")
              If Not AdoProv.Recordset.EOF Then
                 DCProv.Text = AdoProv.Recordset.Fields("Cliente")
                 CodigoCliente = AdoProv.Recordset.Fields("Codigo")
              Else
                 DCProv.Text = "PROVEEDOR NO ASIGNADO"
              End If
           End If
           If AdoResp.Recordset.RecordCount > 0 Then
              AdoResp.Recordset.MoveFirst
              AdoResp.Recordset.Find ("Codigo = '" & .Fields("Responsable") & "' ")
              If Not AdoResp.Recordset.EOF Then
                 DCResp.Text = AdoResp.Recordset.Fields("Cliente")
                 CodigoBenef = AdoResp.Recordset.Fields("Codigo")
              Else
                 DCResp.Text = "RESPONSABLE NO ASIGNADO"
              End If
           End If
        Else
            MsgBox "No existe"
        End If
    Else
        Nuevo = True
        TxtSubCta.SetFocus
    End If
   End With
End Sub

Public Sub GrabarInv()
  RatonReloj
  Nuevo = False
  TextoValido TxtUnidad
  TextoValido TxtBarra
  If Len(TxtDetalle.Text) <= 1 Then TxtDetalle.Text = Ninguno
 'CampoBusqueda = DGBusq.Columns(DGBusq.Col).Caption
  Codigo = UCase(CambioCodigoCta(MBoxCodigo))
  If Option1.value Then TxtSubCta.Text = UCase(TxtSubCta.Text)
  Codigo1 = "C" & Codigo
  Cta_Sup = "C" & CodigoCuentaSup(Codigo)
  Cuenta = Codigo & " - " & TxtSubCta.Text
  If Option1.value Then Cadena = "I" Else Cadena = "P"
  TipoCta = Cadena
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo_Activo "
  SelectAdodc AdoInv, sSQL
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       TextoBusqueda = "Codigo_Activo Like '" & Codigo & "' "
      .Find (TextoBusqueda)
       If .EOF Then
           SetAddNew AdoInv
           Nuevo = False
       End If
   Else
      SetAddNew AdoInv
      Nuevo = True
   End If
   'MsgBox Nuevo & vbCrLf & Codigo
   SetFields AdoInv, "Cta_Activo", "0"
   SetFields AdoInv, "Codigo_Activo", Codigo
   SetFields AdoInv, "Producto", TxtSubCta
   SetFields AdoInv, "TC", TipoCta
   SetFields AdoInv, "Ubicacion", TxtUnidad
   SetFields AdoInv, "Codigo_Barra", TxtBarra
   SetFields AdoInv, "Detalle", TxtDetalle
   SetFields AdoInv, "Sub_Total", TxtSubTotal
   SetFields AdoInv, "Total_IVA", TxtIVA
   SetFields AdoInv, "Total_Compra", LblTotal.Caption
   SetFields AdoInv, "Cantidad", TxtCantidad
   SetFields AdoInv, "Fecha_Compra", MBFecha
   SetFields AdoInv, "Factura_No", TxtFactura
   SetFields AdoInv, "Codigo_Prov", CodigoCliente
   SetFields AdoInv, "Responsable", CodigoBenef
   If TipoCta <> "I" Then
      SetFields AdoInv, "Cta_Activo", CambioCodigoCta(MBoxCta_Inv)
      SetFields AdoInv, "Cta_Patrimonio", CambioCodigoCta(MBoxCta_Pat)
   End If
   If CheqConta.value = 1 Then SetFields AdoInv, "T", "C" Else SetFields AdoInv, "T", "N"
   SetFields AdoInv, "Periodo", Periodo_Contable
   SetFields AdoInv, "Item", NumEmpresa
   SetUpdate AdoInv
   If Nuevo Then
      Codigo2 = Codigo
      Codigo = Codigo1
      AddNewCtaInv TipoCta
      Codigo = Codigo2
   Else
      IE = TVCatalogo.SelectedItem.Index
      TVCatalogo.Nodes(IE).Text = Codigo & " - " & TxtSubCta.Text
      TVCatalogo.Refresh
   End If
  End With
  RatonNormal
  MsgBox "Proceso exitoso"
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyDelete Then
     Codigo = SinEspaciosIzq(TVCatalogo.SelectedItem)
     Cuenta = SinEspaciosDer(TVCatalogo.SelectedItem)
     Mensajes = "Seguro de Eliminar el Codigo:" & Codigo & vbCrLf & "?"
     Titulo = "ELIMINACION"
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * " _
             & "FROM Catalogo_Productos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo_Activo = '" & Codigo & "' "
        ConectarAdoExecute sSQL
        TVCatalogo.Nodes.Remove TVCatalogo.SelectedItem.Index
    End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then ImprimirAdodc AdoInv, True, 2, 8
  If CtrlDown And KeyCode = vbKeyU Then
     Unload IngActivos
  End If
End Sub

Private Sub TVCatalogo_LostFocus()
  LlenarInv
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTexto TxtDetalle
End Sub

Private Sub TxtSubTotal_GotFocus()
  MarcarTexto TxtSubTotal
End Sub

Private Sub TxtSubTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtUnidad_GotFocus()
  MarcarTexto TxtUnidad
End Sub

Private Sub TxtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
