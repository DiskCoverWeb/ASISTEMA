VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FRetencion 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   FillColor       =   &H00FFC0C0&
   Icon            =   "FRetencion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MouseIcon       =   "FRetencion.frx":030A
   ScaleHeight     =   8274.749
   ScaleMode       =   0  'User
   ScaleWidth      =   11976.19
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtSalarioBasico 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   105
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Region:"
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
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2220
      Begin VB.OptionButton OpcCosta 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Costa"
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
         Left            =   1155
         TabIndex        =   2
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton OpcSierra 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sierra"
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
         TabIndex        =   1
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Importar Rol Pagos"
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
      Left            =   10710
      Picture         =   "FRetencion.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1155
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6210
      Left            =   105
      TabIndex        =   12
      Top             =   1365
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   876
      BackColor       =   12648447
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- Liquidacion de Impuesto"
      TabPicture(0)   =   "FRetencion.frx":0E96
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label24"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label26"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label27"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label20"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label22"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LblSubtotal"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtIngLiqui"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtOtrosIng"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Txt10Tercero"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Txt10Cuarto"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtFondoReserva"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtUtilidades"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtDesahucio"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtApoIESS"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtVivienda"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtSalud"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtEducacion"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtAlimentacion"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtVestimenta"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtDiscap"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Txt3Edad"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtIREmpleador"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtAportePersonal"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "&2.- Consolidacion de Ingresos"
      TabPicture(1)   =   "FRetencion.frx":0EB2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtNumComp"
      Tab(1).Control(1)=   "TxtImpAnterior"
      Tab(1).Control(2)=   "TxtValorRet"
      Tab(1).Control(3)=   "TxtIRCausado"
      Tab(1).Control(4)=   "TxtBaseImpo"
      Tab(1).Control(5)=   "TxtRebajas"
      Tab(1).Control(6)=   "TxtDeduccion"
      Tab(1).Control(7)=   "TxtIngOtrosEmp"
      Tab(1).Control(8)=   "DGConcepto"
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(10)=   "Label30"
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(12)=   "Label7"
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(14)=   "Label29"
      Tab(1).Control(15)=   "Label28"
      Tab(1).Control(16)=   "Label15"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "&3.- Resumen del Rol de Pagos"
      TabPicture(2)   =   "FRetencion.frx":0ECE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DGRolAnio"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4.- Rubros de Todo el año"
      TabPicture(3)   =   "FRetencion.frx":0EEA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGRolPagos"
      Tab(3).ControlCount=   1
      Begin VB.TextBox TxtAportePersonal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   10
         TabIndex        =   28
         Text            =   "9.35"
         Top             =   2880
         Width           =   645
      End
      Begin VB.TextBox TxtNumComp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   63
         Text            =   "0000000"
         Top             =   2880
         Width           =   1800
      End
      Begin VB.TextBox TxtImpAnterior 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   61
         Text            =   "0.00"
         Top             =   2565
         Width           =   1800
      End
      Begin VB.TextBox TxtValorRet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   59
         Text            =   "0.00"
         Top             =   2250
         Width           =   1800
      End
      Begin VB.TextBox TxtIRCausado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "0.00"
         Top             =   1935
         Width           =   1800
      End
      Begin VB.TextBox TxtBaseImpo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   1620
         Width           =   1800
      End
      Begin VB.TextBox TxtRebajas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   53
         Text            =   "0.00"
         Top             =   1305
         Width           =   1800
      End
      Begin VB.TextBox TxtIREmpleador 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   5400
         Width           =   1800
      End
      Begin VB.TextBox Txt3Edad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   5085
         Width           =   1800
      End
      Begin VB.TextBox TxtDiscap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   4770
         Width           =   1800
      End
      Begin VB.TextBox TxtVestimenta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   4455
         Width           =   1800
      End
      Begin VB.TextBox TxtAlimentacion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   4140
         Width           =   1800
      End
      Begin VB.TextBox TxtEducacion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   3825
         Width           =   1800
      End
      Begin VB.TextBox TxtSalud 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   3510
         Width           =   1800
      End
      Begin VB.TextBox TxtVivienda 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3195
         Width           =   1800
      End
      Begin VB.TextBox TxtApoIESS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   2880
         Width           =   1800
      End
      Begin VB.TextBox TxtDesahucio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   2565
         Width           =   1800
      End
      Begin VB.TextBox TxtUtilidades 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   2250
         Width           =   1800
      End
      Begin VB.TextBox TxtFondoReserva 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1935
         Width           =   1800
      End
      Begin VB.TextBox Txt10Cuarto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1620
         Width           =   1800
      End
      Begin VB.TextBox Txt10Tercero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   1305
         Width           =   1800
      End
      Begin VB.TextBox TxtOtrosIng 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   990
         Width           =   1800
      End
      Begin VB.TextBox TxtDeduccion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   990
         Width           =   1800
      End
      Begin VB.TextBox TxtIngOtrosEmp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   -66810
         MaxLength       =   10
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   675
         Width           =   1800
      End
      Begin VB.TextBox TxtIngLiqui 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   675
         Width           =   1800
      End
      Begin MSDataGridLib.DataGrid DGRolPagos 
         Bindings        =   "FRetencion.frx":0F06
         Height          =   5475
         Left            =   -74895
         TabIndex        =   67
         Top             =   630
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   9657
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
      Begin MSDataGridLib.DataGrid DGRolAnio 
         Bindings        =   "FRetencion.frx":0F20
         Height          =   5475
         Left            =   -74895
         TabIndex        =   68
         Top             =   630
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   9657
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
      Begin MSDataGridLib.DataGrid DGConcepto 
         Bindings        =   "FRetencion.frx":0F39
         Height          =   2640
         Left            =   -74580
         TabIndex        =   69
         Top             =   3360
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   4657
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
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número de Retenciones"
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
         Left            =   -74580
         TabIndex        =   62
         ToolTipText     =   "Número de comprobante de Retención"
         Top             =   2865
         Width           =   7785
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor del Impuesto Retenido por Empleadores anteriores durante el período"
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
         Left            =   -74580
         TabIndex        =   60
         Top             =   2550
         Width           =   7785
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor del Impuesto Retenido por este Empleador"
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
         Left            =   -74580
         TabIndex        =   58
         Top             =   2235
         Width           =   7785
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Impuesto a la Renta Causado"
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
         Left            =   -74580
         TabIndex        =   56
         ToolTipText     =   "Base para Retencíon"
         Top             =   1920
         Width           =   7785
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Base Imponible Rentas Total anual"
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
         Left            =   -74580
         TabIndex        =   54
         ToolTipText     =   "Base para Retencíon"
         Top             =   1605
         Width           =   7785
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Otras rebajas consideradas por otros Empleadores"
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
         Left            =   -74580
         TabIndex        =   52
         Top             =   1290
         Width           =   7785
      End
      Begin VB.Label LblSubtotal 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   8190
         TabIndex        =   47
         ToolTipText     =   "Serie Comprobante de Retención"
         Top             =   5700
         Width           =   1800
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " S U B T O T A L"
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
         Left            =   420
         TabIndex        =   46
         ToolTipText     =   "Serie Comprobante de Retención"
         Top             =   5700
         Width           =   7785
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Impuesto a la Renta asumido por el empleador"
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
         Left            =   420
         TabIndex        =   44
         Top             =   5385
         Width           =   7785
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Rebajas Especiales Tercera Edad  "
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
         Left            =   420
         TabIndex        =   42
         Top             =   5070
         Width           =   7785
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Rebajas Especiales Discapacitados  "
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
         Left            =   420
         TabIndex        =   40
         Top             =   4740
         Width           =   7785
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción Gastos Personales - Vestimenta"
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
         Left            =   420
         TabIndex        =   38
         Top             =   4425
         Width           =   7785
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción Gastos Personales - Alimentación"
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
         Left            =   420
         TabIndex        =   36
         Top             =   4125
         Width           =   7785
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción Gastos Personales - Educación"
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
         Left            =   420
         TabIndex        =   34
         Top             =   3810
         Width           =   7785
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción Gastos Personales - Salud"
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
         Left            =   420
         TabIndex        =   32
         Top             =   3495
         Width           =   7785
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción Gastos Personales - Vivienda"
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
         Left            =   420
         TabIndex        =   30
         Top             =   3180
         Width           =   7785
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Aporte Personal IESS (únicamente pagado por el empleado)                        (%)"
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
         Left            =   420
         TabIndex        =   27
         Top             =   2865
         Width           =   7785
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desahucio y otras remuneraciones que no constituyen Renta gravada (informativo)"
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
         Left            =   420
         TabIndex        =   25
         Top             =   2550
         Width           =   7785
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Participación de utilidades"
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
         Left            =   420
         TabIndex        =   23
         Top             =   2235
         Width           =   7785
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fondo de Reserva (informativo)"
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
         Left            =   420
         TabIndex        =   21
         Top             =   1920
         Width           =   7785
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Décimo cuarto sueldo (informativo)"
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
         Left            =   420
         TabIndex        =   19
         Top             =   1590
         Width           =   7785
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Décimo tercer sueldo (informativo)"
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
         Left            =   420
         TabIndex        =   17
         Top             =   1275
         Width           =   7785
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sobre Sueldos, comisiones y otras remunerac."
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
         Left            =   420
         TabIndex        =   15
         Top             =   975
         Width           =   7785
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Deducción gastos personales consideradas por otros Empleadores"
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
         Left            =   -74580
         TabIndex        =   50
         ToolTipText     =   "Base para Retencíon"
         Top             =   975
         Width           =   7785
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ingresos gravados generados por otros Empleadores"
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
         Left            =   -74580
         TabIndex        =   48
         ToolTipText     =   "Base para Retencíon"
         Top             =   660
         Width           =   7785
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sueldos y Salarios"
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
         Left            =   420
         TabIndex        =   13
         Top             =   660
         Width           =   7785
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Grabar Anexo"
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
      Left            =   10710
      Picture         =   "FRetencion.frx":0F53
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   105
      Width           =   1065
   End
   Begin VB.TextBox TxtNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   9660
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "0"
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Salir Módulo"
      DisabledPicture =   "FRetencion.frx":1395
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
      Left            =   10710
      MouseIcon       =   "FRetencion.frx":1DDF
      Picture         =   "FRetencion.frx":2829
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2205
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo DCRetenido 
      Bindings        =   "FRetencion.frx":2C6B
      DataSource      =   "AdoRetenido"
      Height          =   315
      Left            =   3885
      TabIndex        =   8
      Top             =   525
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc AdoRetenido 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "Retenido"
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
   Begin MSAdodcLib.Adodc AdoRetencion 
      Height          =   330
      Left            =   210
      Top             =   2625
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
      Caption         =   "Retencion"
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   210
      Top             =   3570
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
      Caption         =   "DetRet"
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
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   9240
      TabIndex        =   6
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   0
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
   Begin MSAdodcLib.Adodc AdoConcepto 
      Height          =   330
      Left            =   210
      Top             =   2940
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
      Caption         =   "Concepto"
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
   Begin MSAdodcLib.Adodc AdoQuery1 
      Height          =   330
      Left            =   210
      Top             =   3255
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
      Caption         =   "PorcDep"
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
   Begin MSAdodcLib.Adodc AdoRolPagos 
      Height          =   330
      Left            =   210
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
      Caption         =   "RolPagos"
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
   Begin MSAdodcLib.Adodc AdoRolAnio 
      Height          =   330
      Left            =   210
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
      Caption         =   "RolAnio"
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
      Left            =   210
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
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Salario &Básico:"
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
      Left            =   2415
      TabIndex        =   3
      ToolTipText     =   "Retenido a:"
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sistema de Salario Neto"
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
      TabIndex        =   9
      ToolTipText     =   "Retenido a:"
      Top             =   945
      Width           =   5265
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Numero de mese trabajados con este Empleado"
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
      TabIndex        =   10
      ToolTipText     =   "Número de comprobante Contable"
      Top             =   945
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Empleado:"
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
      Left            =   2415
      TabIndex        =   7
      ToolTipText     =   "Retenido a:"
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha de Generación del RDEP (107):"
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
      Height          =   360
      Left            =   5775
      TabIndex        =   5
      ToolTipText     =   "Informacion para el mes de:"
      Top             =   105
      Width           =   3495
   End
End
Attribute VB_Name = "FRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FechaDec3Ini As String
Dim FechaDec3Fin As String
Dim FechaDec4Ini As String
Dim FechaDec4Fin As String

Public Sub Encerar_Datos_Dependencia()
   TxtIngLiqui = "0.00"
   TxtOtrosIng = "0.00"
   Txt10Tercero = "0.00"
   Txt10Cuarto = "0.00"
   TxtUtilidades = "0.00"
   TxtApoIESS = "0.00"
   TxtDiscap = "0.00"
   Txt3Edad = "0.00"
   LblSubtotal.Caption = "0.00"
   TxtIREmpleador = "0.00"
   TxtIRCausado = "0.00"
   TxtBaseImpo = "0.00"
   TxtValorRet = "0.00"
   TxtNumComp = "0.00"
   TxtNumero = "0"
   TxtVivienda = "0.00"
   TxtSalud = "0.00"
   TxtEducacion = "0.00"
   TxtAlimentacion = "0.00"
   TxtVestimenta = "0.00"
   TxtDiscap = "0.00"
   TxtIngOtrosEmp = "0.00"
   TxtDeduccion = "0.00"
   TxtRebajas = "0.00"
   TxtImpAnterior = "0.00"
   TxtFondoReserva = "0.00"
   TxtDesahucio = "0.00"
   TxtAportePersonal = "0.00"
End Sub

Public Sub Grabar_Anexo_Dependencia(CodigoC As String)
    RatonReloj
    With AdoQuery1.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Codigo = '" & CodigoC & "' ")
         If Not .EOF Then
           'If CodigoC = "0200670800" Then MsgBox "..."
           .Fields("SuelSal") = Val(CCur(TxtIngLiqui))
           .Fields("SobSuelComRemu") = Val(CCur(TxtOtrosIng))
           .Fields("Valor_Dec_3ro") = Val(CCur(Txt10Tercero))
           .Fields("Valor_Dec_4to") = Val(CCur(Txt10Cuarto))
           .Fields("PartUtil") = Val(CCur(TxtUtilidades))
           .Fields("ApoPerIess") = Val(CCur(TxtApoIESS))
           .Fields("RebEspDiscap") = Val(CCur(TxtDiscap))
           .Fields("RebEspTerEd") = Val(CCur(Txt3Edad))
           .Fields("SubTotal") = Val(CCur(LblSubtotal.Caption))
           .Fields("ImpRentEmpl") = Val(CCur(TxtIREmpleador))
           .Fields("ImpRentCaus") = Val(CCur(TxtIRCausado))
           .Fields("BasImp") = Val(CCur(TxtBaseImpo))
           .Fields("ValRet") = Val(CCur(TxtValorRet))
           .Fields("NumRet") = Val(CInt(TxtNumComp))
           .Fields("AñoRet") = Year(MBoxFechaI)
           .Fields("Fecha_107") = MBoxFechaI
           .Fields("Numero") = Val(CInt(TxtNumero))
           .Fields("Vivienda") = Val(CCur(TxtVivienda))
           .Fields("Salud") = Val(CCur(TxtSalud))
           .Fields("Educacion") = Val(CCur(TxtEducacion))
           .Fields("Alimentacion") = Val(CCur(TxtAlimentacion))
           .Fields("Vestimenta") = Val(CCur(TxtVestimenta))
           .Fields("Discapacidad") = Val(CCur(TxtDiscap))
           .Fields("IngOtrosEmp") = Val(CCur(TxtIngOtrosEmp))
           .Fields("Deduccion") = Val(CCur(TxtDeduccion))
           .Fields("Rebajas") = Val(CCur(TxtRebajas))
           .Fields("ImpAnterior") = Val(CCur(TxtImpAnterior))
           .Fields("FondoReserva") = Val(CCur(TxtFondoReserva))
           .Fields("Desahucio") = Val(CCur(TxtDesahucio))
           .Fields("CodigoU") = CodigoUsuario
           .Fields("Linea_SRI") = 0
           .Update
         End If
     End If
    End With
    RatonNormal
End Sub

Public Function RDep_Subtotal()
Dim Subtotales As Currency
Dim BaseImponible As Currency
Dim ImpuestoCausado As Currency
'''   Txt10Tercero = Format(Val(CCur(TxtIngLiqui)) / 12, "#,##0.00")
'''   Txt10Cuarto = Format((Sueldo_Basico / 12) * Val(CInt(TxtNumero)), "#,##0.00")
'''   If Val(CCur(TxtAportePersonal)) > 0 Then
'''      TxtApoIESS = Format((Val(CCur(TxtIngLiqui)) + Val(CCur(TxtOtrosIng))) * Val(CCur(TxtAportePersonal)) / 100, "#,##0.00")
'''   Else
'''      TxtApoIESS = "0.00"
'''   End If
   Subtotales = 0
   Subtotales = Subtotales + Val(CCur(TxtIngLiqui))
   Subtotales = Subtotales + Val(CCur(TxtOtrosIng))
   Subtotales = Subtotales + Val(CCur(TxtUtilidades))
   Subtotales = Subtotales - Val(CCur(TxtApoIESS))
   Subtotales = Subtotales - Val(CCur(TxtVivienda))
   Subtotales = Subtotales - Val(CCur(TxtSalud))
   Subtotales = Subtotales - Val(CCur(TxtEducacion))
   Subtotales = Subtotales - Val(CCur(TxtAlimentacion))
   Subtotales = Subtotales - Val(CCur(TxtVestimenta))
   Subtotales = Subtotales - Val(CCur(txtTxtDiscap))
   Subtotales = Subtotales - Val(CCur(Txt3Edad))
   Subtotales = Subtotales + Val(CCur(TxtIREmpleador))
   LblSubtotal = Format(Subtotales, "#,##0.00")
   
   BaseImponible = Subtotales
   BaseImponible = BaseImponible + Val(CCur(TxtIngOtrosEmp))
   BaseImponible = BaseImponible - Val(CCur(TxtDeduccion))
   BaseImponible = BaseImponible - Val(CCur(TxtRebajas))
   
   ImpuestoCausado = 0
   With AdoConcepto.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           If (.Fields("Desde") <= BaseImponible) And (BaseImponible <= .Fields("Hasta")) Then
              ImpuestoCausado = BaseImponible - .Fields("Desde")
              ImpuestoCausado = (ImpuestoCausado * .Fields("Excede")) / 100
              ImpuestoCausado = ImpuestoCausado + .Fields("Basico")
             .MoveLast
           End If
          .MoveNext
        Loop
    Else
       MsgBox "No a seleccionado el año fiscal a declarar correctamente"
    End If
  End With
  TxtBaseImpo = Format(BaseImponible, "#,##0.00")
  TxtIRCausado = Format(ImpuestoCausado, "#,##0.00")
End Function

Private Sub Command1_Click()
  Unload Me  ' Cancelar
End Sub

Private Sub Command2_Click()
Dim Salario_Neto As Byte
Dim Salario_Basico As Currency
   If ClaveAuxiliar Then
      Salario_Basico = Val(TxtSalarioBasico)
      Mensajes = "Usted  al  procesar  esta  opción," & vbCrLf & vbCrLf _
               & "volvera a importar los datos del Rol" & vbCrLf & vbCrLf _
               & "al Anexo por Relación de Dependencia" & vbCrLf & vbCrLf _
               & "y todos los cambios que ha realizado" & vbCrLf & vbCrLf _
               & "se perderan y tendrá que volver a procesar" & vbCrLf & vbCrLf _
               & "Esta seguro de importar nuevamente?" & vbCrLf & vbCrLf
      If BoxMensaje = vbYes Then
        'Empezamos a ingresar los datos de los empleados
         Encerar_Datos_Dependencia
         With AdoRolPagos.Recordset
          If .RecordCount > 0 Then
              CodigoCliente = .Fields("Codigo")
              If AdoRetenido.Recordset.RecordCount > 0 Then
                 AdoRetenido.Recordset.MoveFirst
                 AdoRetenido.Recordset.Find ("Codigo = '" & CodigoCliente & "' ")
                 If Not AdoRetenido.Recordset.EOF Then
                    'IESS_Per = AdoRetenido.Recordset.Fields("IESS_Per")
                    Salario_Neto = AdoRetenido.Recordset.Fields("SN")
                    If Salario_Neto = 1 Then Label3.Caption = "Sin Sistema de Salario Neto" Else Label3.Caption = "Con Sistema de Salario Neto"
                    TxtAportePersonal = IESS_Per * 100
                    TxtNumero = .Fields("No_Meses")
                 End If
              End If
              Do While Not .EOF
                 If CodigoCliente <> .Fields("Codigo") Then
                   'Importamos lo grabado del rol
                    TxtOtrosIng = "0.00"
                    sSQL = "SELECT TRP.Codigo, CC.DG, SUM(TRP.Ingresos) As Total_Ing " _
                         & "FROM Trans_Rol_de_Pagos As TRP, Catalogo_Cuentas As CC " _
                         & "WHERE TRP.Item = '" & NumEmpresa & "' " _
                         & "AND TRP.Periodo = '" & Periodo_Contable & "' " _
                         & "AND TRP.Codigo = '" & CodigoCliente & "' " _
                         & "AND TRP.Ingresos > 0 " _
                         & "AND CC.Con_IESS <> " & Val(adFalse) & " " _
                         & "AND TRP.Item = CC.Item " _
                         & "AND TRP.Periodo = CC.Periodo " _
                         & "AND TRP.Cta = CC.Codigo " _
                         & "GROUP BY TRP.Codigo, CC.DG "
                    Select_Adodc AdoAux, sSQL
                    If AdoAux.Recordset.RecordCount > 0 Then TxtOtrosIng = AdoAux.Recordset.Fields("Total_Ing")
                    RDep_Subtotal
                    TxtFondoReserva = Format(Val(CCur(TxtIngLiqui)) / 12, "#,##0.00")
                    'Txt10Tercero = Format(Val(CCur(TxtIngLiqui)) / 12, "#,##0.00")
                    Txt10Cuarto = Format((Salario_Basico / 12) * Val(CInt(TxtNumero)), "#,##0.00")
                    Grabar_Anexo_Dependencia CodigoCliente
                    Encerar_Datos_Dependencia
                    CodigoCliente = .Fields("Codigo")
                    If AdoRetenido.Recordset.RecordCount > 0 Then
                       AdoRetenido.Recordset.MoveFirst
                       AdoRetenido.Recordset.Find ("Codigo = '" & CodigoCliente & "' ")
                       If Not AdoRetenido.Recordset.EOF Then
                          'IESS_Per = AdoRetenido.Recordset.Fields("IESS_Per")
                          Salario_Neto = AdoRetenido.Recordset.Fields("SN")
                          TxtAportePersonal = IESS_Per * 100
                          TxtNumero = .Fields("No_Meses")
                       End If
                    End If
                 End If
                'If .Fields("Codigo") = "1312041773" And .Fields("Cod_Rol_Pago") = "Decimo_IV" Then MsgBox .Fields("Total_Egr")
                'Escojo los datos del empleado actual
                 Select Case .Fields("Cod_Rol_Pago")
                   Case "Salario": If .Fields("Total_Ing") > 0 Then TxtIngLiqui = .Fields("Total_Ing")
                   Case "Decimo_III": If .Fields("Total_Egr") > 0 Then Txt10Tercero = .Fields("Total_Egr")
                   Case "Decimo_IV": If .Fields("Total_Egr") > 0 Then Txt10Cuarto = .Fields("Total_Egr")
                   Case "Fon_Res_G":   If .Fields("Total_Ing") > 0 Then TxtFondoReserva = .Fields("Total_Ing")
                   Case "Aporte_Per":  If .Fields("Total_Egr") > 0 Then TxtApoIESS = .Fields("Total_Egr")
                   Case "Aporte_Pat":  If .Fields("Total_Egr") > 0 Then TxtApoIESS = .Fields("Total_Egr")
                  'Case "Neto_Recibir":  If .Fields("No_Meses") > 0 Then TxtNumero = .Fields("No_Meses")
                 End Select
                .MoveNext
              Loop
              TxtOtrosIng = "0.00"
              sSQL = "SELECT TRP.Codigo, CC.DG, SUM(TRP.Ingresos) As Total_Ing " _
                   & "FROM Trans_Rol_de_Pagos As TRP, Catalogo_Cuentas As CC " _
                   & "WHERE TRP.Item = '" & NumEmpresa & "' " _
                   & "AND TRP.Periodo = '" & Periodo_Contable & "' " _
                   & "AND TRP.Codigo = '" & CodigoCliente & "' " _
                   & "AND TRP.Ingresos > 0 " _
                   & "AND CC.Con_IESS <> " & Val(adFalse) & " " _
                   & "AND TRP.Item = CC.Item " _
                   & "AND TRP.Periodo = CC.Periodo " _
                   & "AND TRP.Cta = CC.Codigo " _
                   & "GROUP BY TRP.Codigo, CC.DG "
              Select_Adodc AdoAux, sSQL
              If AdoAux.Recordset.RecordCount > 0 Then TxtOtrosIng = AdoAux.Recordset.Fields("Total_Ing")
              RDep_Subtotal
              Grabar_Anexo_Dependencia CodigoCliente
             .MoveFirst
          End If
         End With
      End If
   End If
End Sub

Public Sub Presentar_Rol_Anual_Empleado(Optional CodigoEmpleado As String)
  If CodigoEmpleado = "" Then CodigoEmpleado = Ninguno
  sSQL = ""
  SQL2 = "SELECT C.Cliente As Empleado,TRP.Detalle,SUM(TRP.Ingresos) As Total_Ing,SUM(TRP.Egresos) AS Total_Egr," _
       & "TRP.Fecha_D,TRP.Fecha_H,COUNT(TRP.ID) As No_Meses,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Codigo " _
       & "FROM Trans_Rol_de_Pagos As TRP,Clientes As C " _
       & "WHERE TRP.Item = '" & NumEmpresa & "' "
  SQL1 = SQL2 _
       & "AND TRP.Fecha_D >= #" & FechaIni & "# " _
       & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
       & "AND TRP.Cod_Rol_Pago IN ('Salario','Fon_Res_G') " _
       & "AND TRP.Ingresos > 0 " _
       & "AND TRP.Codigo = '" & CodigoEmpleado & "' " _
       & "AND TRP.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente,TRP.Codigo,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,TRP.Fecha_D,TRP.Fecha_H "
  sSQL = sSQL & SQL1
  SQL1 = SQL2 _
       & "AND TRP.Fecha_D >= #" & FechaIni & "# " _
       & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
       & "AND TRP.Cod_Rol_Pago IN ('Aporte_Pat','Aporte_Per') " _
       & "AND TRP.Egresos > 0 " _
       & "AND TRP.Codigo = '" & CodigoEmpleado & "' " _
       & "AND TRP.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente,TRP.Codigo,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,TRP.Fecha_D,TRP.Fecha_H "
  sSQL = sSQL & "UNION " & SQL1
  SQL1 = SQL2 _
       & "AND TRP.Fecha_D >= #" & FechaDec4Ini & "# " _
       & "AND TRP.Fecha_H <= #" & FechaDec4Fin & "# " _
       & "AND TRP.Cod_Rol_Pago = 'Decimo_IV' " _
       & "AND TRP.Egresos > 0 " _
       & "AND TRP.Codigo = '" & CodigoEmpleado & "' " _
       & "AND TRP.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente,TRP.Codigo,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,TRP.Fecha_D,TRP.Fecha_H "
  sSQL = sSQL & "UNION " & SQL1
  SQL1 = SQL2 _
       & "AND TRP.Fecha_D >= #" & FechaDec3Ini & "# " _
       & "AND TRP.Fecha_H <= #" & FechaDec3Fin & "# " _
       & "AND TRP.Cod_Rol_Pago = 'Decimo_III' " _
       & "AND TRP.Egresos > 0 " _
       & "AND TRP.Codigo = '" & CodigoEmpleado & "' " _
       & "AND TRP.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente,TRP.Codigo,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,TRP.Fecha_D,TRP.Fecha_H "
  sSQL = sSQL & "UNION " & SQL1
  sSQL = sSQL & "ORDER BY Empleado,TRP.Detalle DESC,TRP.Fecha_D "
'  MsgBox sSQL
  Select_Adodc_Grid DGRolAnio, AdoRolAnio, sSQL, 2
End Sub

Private Sub Command3_Click()
  Mensajes = "Desea Grabar la Transacción S/N? "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then Grabar_Anexo_Dependencia CodigoCliente
  Encerar_Datos_Dependencia
  Anio = Format(Year(MBoxFechaI), "0000")
  'Listamos el catalogo del rol de pagos
   SQL2 = "SELECT * " _
        & "FROM Catalogo_Rol_Pagos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Codigo "
   Select_Adodc AdoQuery1, SQL2
  DCRetenido.SetFocus
  'MBoxFechaI.SetFocus
End Sub

Private Sub DCRetenido_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF Then
     With AdoQuery1.Recordset
      If .RecordCount Then
          RatonReloj
         .MoveFirst
          Do While Not .EOF
             TxtIngLiqui = .Fields("SuelSal")
             TxtOtrosIng = .Fields("SobSuelComRemu")
             Txt10Tercero = .Fields("Valor_Dec_3ro")
             Txt10Cuarto = .Fields("Valor_Dec_4to")
             TxtUtilidades = .Fields("PartUtil")
             TxtFondoReserva = .Fields("FondoReserva")
             TxtDesahucio = .Fields("Desahucio")
             TxtApoIESS = .Fields("ApoPerIess")
             TxtDiscap = .Fields("RebEspDiscap")
             Txt3Edad = .Fields("RebEspTerEd")
             LblSubtotal.Caption = .Fields("SubTotal")
             TxtIREmpleador = .Fields("ImpRentEmpl")
             TxtIRCausado = .Fields("ImpRentCaus")
             TxtBaseImpo = .Fields("BasImp")
             TxtValorRet = .Fields("ValRet")
             TxtNumComp = .Fields("NumRet")
             TxtNumero = .Fields("Numero")
             TxtVivienda = .Fields("Vivienda")
             TxtSalud = .Fields("Salud")
             TxtEducacion = .Fields("Educacion")
             TxtAlimentacion = .Fields("Alimentacion")
             TxtVestimenta = .Fields("Vestimenta")
             TxtDiscap = .Fields("Discapacidad")
             TxtIngOtrosEmp = .Fields("IngOtrosEmp")
             TxtDeduccion = .Fields("Deduccion")
             TxtRebajas = .Fields("Rebajas")
             TxtImpAnterior = .Fields("ImpAnterior")
             TxtAportePersonal = .Fields("IESS_Per") * 100
            'Recalculamos Subtotal
             RDep_Subtotal
            'Vovemos a grabar
            .Fields("SuelSal") = Val(CCur(TxtIngLiqui))
            .Fields("SobSuelComRemu") = Val(CCur(TxtOtrosIng))
            .Fields("Valor_Dec_3ro") = Val(CCur(Txt10Tercero))
            .Fields("Valor_Dec_4to") = Val(CCur(Txt10Cuarto))
            .Fields("PartUtil") = Val(CCur(TxtUtilidades))
            .Fields("FondoReserva") = Val(CCur(TxtFondoReserva))
            .Fields("Desahucio") = Val(CCur(TxtDesahucio))
            .Fields("ApoPerIess") = Val(CCur(TxtApoIESS))
            .Fields("RebEspDiscap") = Val(CCur(TxtDiscap))
            .Fields("RebEspTerEd") = Val(CCur(Txt3Edad))
            .Fields("SubTotal") = Val(CCur(LblSubtotal.Caption))
            .Fields("ImpRentEmpl") = Val(CCur(TxtIREmpleador))
            .Fields("ImpRentCaus") = Val(CCur(TxtIRCausado))
            .Fields("BasImp") = Val(CCur(TxtBaseImpo))
            .Fields("ValRet") = Val(CCur(TxtValorRet))
            .Fields("NumRet") = Val(CCur(TxtNumComp))
            .Fields("Numero") = Val(CInt(TxtNumero))
            .Fields("Vivienda") = Val(CCur(TxtVivienda))
            .Fields("Salud") = Val(CCur(TxtSalud))
            .Fields("Educacion") = Val(CCur(TxtEducacion))
            .Fields("Alimentacion") = Val(CCur(TxtAlimentacion))
            .Fields("Vestimenta") = Val(CCur(TxtVestimenta))
            .Fields("Discapacidad") = Val(CCur(TxtDiscap))
            .Fields("IngOtrosEmp") = Val(CCur(TxtIngOtrosEmp))
            .Fields("Deduccion") = Val(CCur(TxtDeduccion))
            .Fields("Rebajas") = Val(CCur(TxtRebajas))
            .Fields("ImpAnterior") = Val(CCur(TxtImpAnterior))
            .Update
            .MoveNext
          Loop
         .MoveFirst
          RatonNormal
        End If
      End With
  End If
  PresionoEnter KeyCode
End Sub

Private Sub DCRetenido_LostFocus()
  FechaValida MBoxFechaI
  Anio = Format(Year(MBoxFechaI), "0000")
  FechaIni = BuscarFecha("01/01/" & Anio)
  FechaFin = BuscarFecha(MBoxFechaI)
  Cadena2 = DCRetenido
  CodigoCliente = Ninguno
  Encerar_Datos_Dependencia
  TxtAportePersonal = "0.00"
  With AdoRetenido.Recordset
   If .RecordCount Then
      .MoveFirst
      .Find ("Cliente = '" & Cadena2 & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          If .Fields("SN") = 1 Then TxtAportePersonal = IESS_Per * 100 Else TxtAportePersonal = "0.00"
       Else
          MsgBox "No existe"
       End If
   End If
  End With
 'Buscamos datos ingresados anteriormente
  
  With AdoQuery1.Recordset
   If .RecordCount Then
      .MoveFirst
      .Find ("Codigo = '" & CodigoCliente & "' ")
       If Not .EOF Then
          If .Fields("SN") = 2 Then Label3.Caption = "Sin Sistema de Salario Neto" Else Label3.Caption = "Con Sistema de Salario Neto"
          TxtNumero = .Fields("Numero")
          TxtIngLiqui = Format(.Fields("SuelSal"), "#,##0.00")
          TxtOtrosIng = Format(.Fields("SobSuelComRemu"), "#,##0.00")
          Txt10Tercero = Format(.Fields("Valor_Dec_3ro"), "#,##0.00")
          Txt10Cuarto = Format(.Fields("Valor_Dec_4to"), "#,##0.00")
          TxtUtilidades = Format(.Fields("PartUtil"), "#,##0.00")
          TxtFondoReserva = Format(.Fields("FondoReserva"), "#,##0.00")
          TxtDesahucio = Format(.Fields("Desahucio"), "#,##0.00")
          TxtApoIESS = Format(.Fields("ApoPerIess"), "#,##0.00")
          TxtDiscap = Format(.Fields("RebEspDiscap"), "#,##0.00")
          Txt3Edad = Format(.Fields("RebEspTerEd"), "#,##0.00")
          LblSubtotal.Caption = Format(.Fields("SubTotal"), "#,##0.00")
          TxtIREmpleador = Format(.Fields("ImpRentEmpl"), "#,##0.00")
          TxtIRCausado = Format(.Fields("ImpRentCaus"), "#,##0.00")
          TxtBaseImpo = Format(.Fields("BasImp"), "#,##0.00")
          TxtValorRet = Format(.Fields("ValRet"), "#,##0.00")
          TxtNumComp = Format(.Fields("NumRet"), "#,##0.00")
          TxtVivienda = Format(.Fields("Vivienda"), "#,##0.00")
          TxtSalud = Format(.Fields("Salud"), "#,##0.00")
          TxtEducacion = Format(.Fields("Educacion"), "#,##0.00")
          TxtAlimentacion = Format(.Fields("Alimentacion"), "#,##0.00")
          TxtVestimenta = Format(.Fields("Vestimenta"), "#,##0.00")
          TxtDiscap = Format(.Fields("Discapacidad"), "#,##0.00")
          TxtIngOtrosEmp = Format(.Fields("IngOtrosEmp"), "#,##0.00")
          TxtDeduccion = Format(.Fields("Deduccion"), "#,##0.00")
          TxtRebajas = Format(.Fields("Rebajas"), "#,##0.00")
          TxtImpAnterior = Format(.Fields("ImpAnterior"), "#,##0.00")
       End If
   End If
  End With
  Presentar_Rol_Anual_Empleado CodigoCliente
  If CodigoCliente = Ninguno Then MBoxFechaI.SetFocus Else TxtNumero.SetFocus
  
'''  SQL2 = "SELECT C.Cliente As Empleado,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,SUM(TRP.Ingresos) As Total_Ing,SUM(TRP.Egresos) AS Total_Egr,COUNT(TRP.ID) As No_Meses,TRP.ID,TRP.Codigo,TRP.Fecha_D,TRP.Fecha_H " _
'''       & "FROM Trans_Rol_de_Pagos As TRP,Clientes As C,Catalogo_Rol_Pagos As CRP " _
'''       & "WHERE TRP.Item = '" & NumEmpresa & "' " _
'''       & "AND TRP.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND TRP.Fecha_D >= #" & FechaIni & "# " _
'''       & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
'''       & "AND TRP.Codigo = '" & CodigoCliente & "' " _
'''       & "AND TRP.Codigo = C.Codigo " _
'''       & "AND CRP.Codigo = C.Codigo " _
'''       & "AND CRP.Item = TRP.Item " _
'''       & "AND CRP.Periodo = TRP.Periodo " _
'''       & "GROUP BY C.Cliente,TRP.Codigo,TRP.Cod_Rol_Pago,TRP.Tipo_Rubro,TRP.Detalle,TRP.ID,TRP.Fecha_D,TRP.Fecha_H " _
'''       & "ORDER BY C.Cliente,TRP.ID,TRP.Cod_Rol_Pago,TRP.Detalle "
'''  SelectMSFGrid MSFGRolAnio, AdoRolAnio, SQL2, 2
End Sub

Private Sub DGRolPagos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGRolPagos.Visible = False
     GenerarDataTexto FRetencion, AdoRolPagos
     DGRolPagos.Visible = True
  End If
End Sub

Private Sub Form_Activate()
   MBoxFechaI = "31/12/" & (Year(FechaSistema) - 1)
   sSQL = "UPDATE Trans_Rol_de_Pagos " _
        & "SET Detalle = 'Sueldos' " _
        & "WHERE Detalle = 'Salario Nominal' "
   Ejecutar_SQL_SP sSQL
   FRetencion.Caption = "Formulario de Retención por Relación de Dependencia"
   TxtSalarioBasico = Format(Sueldo_Basico, "#,##0.00")
   Porcen = 0
   ValRet = 0
   AporIess = 0
   IngLiqui = 0
   BaseImpor = 0
   BaseImpon = 0
  'Listamos el catalogo del rol de pagos
   SQL2 = "SELECT * " _
        & "FROM Catalogo_Rol_Pagos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY Codigo "
   Select_Adodc AdoQuery1, SQL2
   
  'Lista de Empleados activos
   sSQL = "SELECT C.TD,C.CI_RUC,C.Cliente,C.Codigo,R.SN " _
        & "FROM Clientes As C, Catalogo_Rol_Pagos As R " _
        & "WHERE R.Codigo = C.Codigo " _
        & "AND R.Item = '" & NumEmpresa & "' " _
        & "AND R.Periodo = '" & Periodo_Contable & "' " _
        & "ORDER BY C.Cliente "
   SelectDB_Combo DCRetenido, AdoRetenido, sSQL, "Cliente"
   RatonNormal
   
End Sub

Private Sub Form_Load()
  'CentrarForm FRetencion
  ConectarAdodc AdoAux
  ConectarAdodc AdoConcepto
  ConectarAdodc AdoRetenido
  ConectarAdodc AdoDetRet
  ConectarAdodc AdoRolPagos
  ConectarAdodc AdoRetencion
  ConectarAdodc AdoQuery1
  ConectarAdodc AdoRolAnio
  Vanio = Year(FechaSistema)
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then: Command1.SetFocus
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI
  Anio = Format(Year(MBoxFechaI), "0000")
  FechaIni = BuscarFecha("01/01/" & Anio)
  FechaFin = BuscarFecha(MBoxFechaI)
  If OpcCosta.value Then
     FechaDec4Ini = BuscarFecha("01/03/" & CStr(Val(Anio)) - 1)
     Mifecha = UltimoDiaMes("01/02/" & Anio)
     FechaDec4Fin = BuscarFecha(Mifecha)
  Else
     FechaDec4Ini = BuscarFecha("01/08/" & CStr(Val(Anio)) - 1)
     Mifecha = UltimoDiaMes("01/07/" & Anio)
     FechaDec4Fin = BuscarFecha(Mifecha)
  End If
  FechaDec3Ini = BuscarFecha("01/12/" & CStr(Val(Anio)) - 1)
  Mifecha = UltimoDiaMes("01/11/" & Anio)
  FechaDec3Fin = BuscarFecha(Mifecha)
  
  sSQL = "UPDATE Trans_Rol_de_Pagos " _
       & "SET Ingresos = ROUND(Ingresos,2,0),Egresos = ROUND(Egresos,2,0) " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL

  SQL2 = "SELECT Desde, Hasta, Basico, Excede " _
       & "FROM Tabla_Renta " _
       & "WHERE Anio = '" & CStr(Anio) & "' " _
       & "ORDER BY Desde,Hasta "
  Select_Adodc_Grid DGConcepto, AdoConcepto, SQL2, 2, True
  Presentar_Rol_Anual
  DCRetenido.SetFocus
End Sub

Private Sub Txt10Cuarto_GotFocus()
  MarcarTexto Txt10Cuarto
End Sub

Private Sub Txt10Cuarto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Txt10Cuarto_LostFocus()
  TextoValido Txt10Cuarto, True, 2
  RDep_Subtotal
End Sub

Private Sub Txt10Tercero_GotFocus()
  MarcarTexto Txt10Tercero
End Sub

Private Sub Txt10Tercero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Txt10Tercero_LostFocus()
  TextoValido Txt10Tercero, True, , 2
  RDep_Subtotal
End Sub

Private Sub Txt3Edad_GotFocus()
  MarcarTexto Txt3Edad
End Sub

Private Sub Txt3Edad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Txt3Edad_LostFocus()
  TextoValido Txt3Edad, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtAlimentacion_GotFocus()
  MarcarTexto TxtAlimentacion
End Sub

Private Sub TxtAlimentacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAlimentacion_LostFocus()
  TextoValido TxtAlimentacion, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtApoIESS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApoIESS_LostFocus()
   TextoValido TxtApoIESS, True, , 2
   RDep_Subtotal
End Sub

Private Sub TxtAportePersonal_GotFocus()
  MarcarTexto TxtAportePersonal
End Sub

Private Sub TxtAportePersonal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAportePersonal_LostFocus()
  TextoValido TxtAportePersonal, True, , 2
  TxtApoIESS = Format((Val(CCur(TxtIngLiqui)) + Val(CCur(TxtOtrosIng))) * Val(CCur(TxtAportePersonal)) / 100, "#,##0.00")
End Sub

Private Sub TxtBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpo_GotFocus()
  MarcarTexto TxtBaseImpo
End Sub

Private Sub TxtBaseImpo_LostFocus()
  TextoValido TxtBaseImpo, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtDeduccion_GotFocus()
  MarcarTexto TxtDeduccion
End Sub

Private Sub TxtDeduccion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDeduccion_LostFocus()
  TextoValido TxtDeduccion, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtDesahucio_GotFocus()
  MarcarTexto TxtDesahucio
End Sub

Private Sub TxtDesahucio_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesahucio_LostFocus()
  TextoValido TxtDesahucio, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtDiscap_GotFocus()
  MarcarTexto TxtDiscap
End Sub

Private Sub TxtDiscap_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDiscap_LostFocus()
  TextoValido TxtDiscap, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtEducacion_GotFocus()
  MarcarTexto TxtEducacion
End Sub

Private Sub TxtEducacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEducacion_LostFocus()
  TextoValido TxtEducacion, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtFondoReserva_GotFocus()
  MarcarTexto TxtFondoReserva
End Sub

Private Sub TxtFondoReserva_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFondoReserva_LostFocus()
  TextoValido TxtFondoReserva, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtImpAnterior_GotFocus()
  MarcarTexto TxtImpAnterior
End Sub

Private Sub TxtImpAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtImpAnterior_LostFocus()
  TextoValido TxtImpAnterior, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtIngOtrosEmp_GotFocus()
  MarcarTexto TxtIngOtrosEmp
End Sub

Private Sub TxtIngOtrosEmp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIngOtrosEmp_LostFocus()
  TextoValido TxtIngOtrosEmp, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtIRCausado_GotFocus()
  MarcarTexto TxtIRCausado
End Sub

Private Sub TxtIRCausado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIRCausado_LostFocus()
  TextoValido TxtIRCausado, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtIREmpleador_GotFocus()
  MarcarTexto TxtIREmpleador
End Sub

Private Sub TxtIREmpleador_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIREmpleador_LostFocus()
  TextoValido TxtIREmpleador, True, , 2
  RDep_Subtotal
  SSTab1.Tab = 1
  TxtIngOtrosEmp.SetFocus
End Sub

Private Sub TxtNumComp_GotFocus()
   If Val(CCur(TxtValorRet)) > 0 Then TxtNumComp = "1" Else TxtNumComp = "0"
   MarcarTexto TxtNumComp
End Sub

Private Sub TxtNumComp_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtNumComp_LostFocus()
   TextoValido TxtNumComp, True, , 0
   RDep_Subtotal
   SSTab1.Tab = 0
   TxtIngLiqui.SetFocus
End Sub

Private Sub TxtApoIESS_GotFocus()
  MarcarTexto TxtApoIESS
End Sub

Private Sub TxtIngLiqui_GotFocus()
  MarcarTexto TxtIngLiqui
End Sub

Private Sub TxtIngLiqui_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIngLiqui_LostFocus()
  TextoValido TxtIngLiqui, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtNumero_GotFocus()
   MarcarTexto TxtNumero
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumero_LostFocus()
   TextoValido TxtNumero, True
   TxtNumero = Format(TxtNumero, "000000")
End Sub

Private Sub TxtOtrosIng_GotFocus()
  MarcarTexto TxtOtrosIng
End Sub

Private Sub TxtOtrosIng_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOtrosIng_LostFocus()
   TextoValido TxtOtrosIng, True, , 2
   RDep_Subtotal
End Sub

Private Sub TxtRebajas_GotFocus()
  MarcarTexto TxtRebajas
End Sub

Private Sub TxtRebajas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRebajas_LostFocus()
  TextoValido TxtRebajas, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtSalarioBasico_GotFocus()
  MarcarTexto TxtSalarioBasico
End Sub

Private Sub TxtSalarioBasico_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSalarioBasico_LostFocus()
  TextoValido TxtSalarioBasico, True, , 2
End Sub

Private Sub TxtSalud_GotFocus()
  MarcarTexto TxtSalud
End Sub

Private Sub TxtSalud_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSalud_LostFocus()
  TextoValido TxtSalud, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtUtilidades_GotFocus()
  MarcarTexto TxtUtilidades
End Sub

Private Sub TxtUtilidades_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtUtilidades_LostFocus()
  TextoValido TxtUtilidades, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtValorRet_GotFocus()
  MarcarTexto TxtValorRet
End Sub

Private Sub TxtValorRet_LostFocus()
  TextoValido TxtValorRet, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtValorRet_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then: Command1.SetFocus
  PresionoEnter KeyCode
End Sub

Private Sub TxtVestimenta_GotFocus()
  MarcarTexto TxtVestimenta
End Sub

Private Sub TxtVestimenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtVestimenta_LostFocus()
  TextoValido TxtVestimenta, True, , 2
  RDep_Subtotal
End Sub

Private Sub TxtVivienda_GotFocus()
  MarcarTexto TxtVivienda
End Sub

Private Sub TxtVivienda_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtVivienda_LostFocus()
  TextoValido TxtVivienda, True, , 2
  RDep_Subtotal
End Sub

Public Sub Presentar_Rol_Anual()
Dim SQLE As String
Dim SQLG As String

  SQLE = "SELECT C.Cliente As Empleado,TRP.Detalle," _
       & "SUM(TRP.Ingresos) As Total_Ing, SUM(TRP.Egresos) AS Total_Egr, COUNT(TRP.ID) As No_Meses, " _
       & "TRP.Cod_Rol_Pago, TRP.Tipo_Rubro, TRP.Codigo " _
       & "FROM Trans_Rol_de_Pagos As TRP, Clientes As C " _
       & "WHERE TRP.Item = '" & NumEmpresa & "' "
       
  SQLG = "AND TRP.Codigo = C.Codigo " _
       & "GROUP BY C.Cliente, TRP.Codigo, TRP.Cod_Rol_Pago, TRP.Tipo_Rubro, TRP.Detalle "
 '======================================================================================================
  sSQL = SQLE _
       & "AND TRP.Fecha_D >= #" & FechaIni & "# " _
       & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
       & "AND TRP.Cod_Rol_Pago IN ('Salario','Fon_Res_G') " _
       & "AND TRP.Ingresos > 0 " _
       & SQLG
  sSQL = sSQL & "UNION "
  sSQL = sSQL & SQLE _
       & "AND TRP.Fecha_D >= #" & FechaIni & "# " _
       & "AND TRP.Fecha_H <= #" & FechaFin & "# " _
       & "AND TRP.Cod_Rol_Pago IN ('Aporte_Pat','Aporte_Per') " _
       & "AND TRP.Egresos > 0 " _
       & SQLG
  sSQL = sSQL & "UNION "
  sSQL = sSQL & SQLE _
       & "AND TRP.Fecha_D >= #" & FechaDec4Ini & "# " _
       & "AND TRP.Fecha_H <= #" & FechaDec4Fin & "# " _
       & "AND TRP.Cod_Rol_Pago = 'Decimo_IV' " _
       & "AND TRP.Egresos > 0 " _
       & SQLG
  sSQL = sSQL & "UNION "
  sSQL = sSQL & SQLE _
       & "AND TRP.Fecha_D >= #" & FechaDec3Ini & "# " _
       & "AND TRP.Fecha_H <= #" & FechaDec3Fin & "# " _
       & "AND TRP.Cod_Rol_Pago = 'Decimo_III' " _
       & "AND TRP.Egresos > 0 " _
       & SQLG
  sSQL = sSQL & "ORDER BY Empleado,Total_Ing DESC,TRP.Detalle,Total_Egr "
 'MsgBox sSQL
  Select_Adodc_Grid DGRolPagos, AdoRolPagos, sSQL, 2
End Sub

