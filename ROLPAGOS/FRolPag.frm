VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form FRolPago 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ROL DE PAGOS"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   Icon            =   "FRolPag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DCGrupos 
      Bindings        =   "FRolPag.frx":014A
      DataSource      =   "AdoGrupos"
      Height          =   315
      Left            =   7455
      TabIndex        =   14
      Top             =   945
      Width           =   2640
      _ExtentX        =   4657
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
   Begin VB.TextBox TxtExtC 
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
      Left            =   9345
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "FRolPag.frx":0162
      ToolTipText     =   "<Ctrl+P> Cambia el Aporte Patronal"
      Top             =   2520
      Width           =   750
   End
   Begin VB.CheckBox CheqExtC 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Ext. Conyugue %"
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
      Left            =   7455
      TabIndex        =   38
      Top             =   2520
      Width           =   1905
   End
   Begin MSMask.MaskEdBox MBFechaM 
      Height          =   330
      Left            =   5985
      TabIndex        =   33
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   2415
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
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
   Begin VB.CheckBox CheqMaternidad 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha de Maternidad"
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
      Left            =   3780
      TabIndex        =   32
      Top             =   2415
      Width           =   2220
   End
   Begin MSMask.MaskEdBox MBFechaC 
      Height          =   330
      Left            =   5985
      TabIndex        =   31
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1995
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
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
   Begin VB.CheckBox CheqSalida 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha de Salida"
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
      Left            =   3780
      TabIndex        =   30
      Top             =   1995
      Width           =   2220
   End
   Begin VB.CheckBox CheqTiempoParcial 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salario Tiempo parcial"
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
      Height          =   225
      Left            =   105
      TabIndex        =   29
      Top             =   2730
      Width           =   2430
   End
   Begin VB.CheckBox CheqB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bachillerato"
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
      Height          =   225
      Left            =   2100
      TabIndex        =   28
      Top             =   2415
      Width           =   1380
   End
   Begin VB.CheckBox CheqS 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Básico"
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
      Height          =   225
      Left            =   1050
      TabIndex        =   27
      Top             =   2415
      Width           =   960
   End
   Begin VB.CheckBox CheqP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Inicial"
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
      Height          =   225
      Left            =   105
      TabIndex        =   26
      Top             =   2415
      Width           =   855
   End
   Begin VB.TextBox TxtPorcCom 
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
      Left            =   5985
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "FRolPag.frx":0166
      Top             =   1575
      Width           =   1380
   End
   Begin VB.TextBox TxtClave 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1995
      MaxLength       =   8
      PasswordChar    =   "¤"
      TabIndex        =   22
      Top             =   1575
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3585
      Left            =   105
      TabIndex        =   40
      Top             =   3045
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   6324
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12648447
      TabCaption(0)   =   "Seteos de Cuentas"
      TabPicture(0)   =   "FRolPag.frx":016A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label20"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label36"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label39"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label33"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label35"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label38"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label41"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label34"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label42"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label53"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label52"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label40"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label37"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label12"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DCSubModulo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "MBCta_ExtConyugue"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "MBCta_Sueldo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "MBCta_IESS_Patronal"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MBCta_Horas_Ext"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "MBCta_Decimo_Tercer_G"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "MBCta_Decimo_Tercer_P"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "MBCta_Vacacion"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "MBCta_IESS_Personal"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "MBCta_Antig"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "MBCta_Decimo_Cuarto_G"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "MBCta_Decimo_Cuarto_P"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "MBCta_Quincena"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "MBCta_Aporte_Patronal_G"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "MBCta_Vacaciones_G"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "MBCta_Vacaciones_P"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "CheqRFR"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "CheqFondoReserva"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "MBCta_Fondo_Reserva_P"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "MBCta_Fondo_Reserva_G"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "MBCta_Diferencia"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "CheqDecimos"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Seguro Social y Decimos"
      TabPicture(1)   =   "FRolPag.frx":0186
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtDiasDec4to"
      Tab(1).Control(1)=   "TxtCISustituye"
      Tab(1).Control(2)=   "TxtPorcDiscap"
      Tab(1).Control(3)=   "TxtDiasDec3ro"
      Tab(1).Control(4)=   "TxtValorDec3ro"
      Tab(1).Control(5)=   "TxtValorDec4to"
      Tab(1).Control(6)=   "TxtCodProfesion"
      Tab(1).Control(7)=   "TxtFPDec"
      Tab(1).Control(8)=   "MBFechaVF"
      Tab(1).Control(9)=   "MBFechaVI"
      Tab(1).Control(10)=   "DCAplicaConvenio"
      Tab(1).Control(11)=   "DCCondiciones"
      Tab(1).Control(12)=   "Label46"
      Tab(1).Control(13)=   "Label48"
      Tab(1).Control(14)=   "Label49"
      Tab(1).Control(15)=   "LblTDsustituye"
      Tab(1).Control(16)=   "Label57"
      Tab(1).Control(17)=   "Label56"
      Tab(1).Control(18)=   "Label55"
      Tab(1).Control(19)=   "Label54"
      Tab(1).Control(20)=   "Label47"
      Tab(1).Control(21)=   "Label44"
      Tab(1).Control(22)=   "Label45"
      Tab(1).Control(23)=   "Label24"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Forma de Pago"
      TabPicture(2)   =   "FRolPag.frx":01A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OpcEfectivo"
      Tab(2).Control(1)=   "OpcCheque"
      Tab(2).Control(2)=   "OpcTransferencia"
      Tab(2).Control(3)=   "OpcOtro"
      Tab(2).Control(4)=   "TxtCI"
      Tab(2).Control(5)=   "TxtCtaAbono"
      Tab(2).Control(6)=   "DCFormaPago"
      Tab(2).Control(7)=   "DCTipoBanco"
      Tab(2).Control(8)=   "Label50"
      Tab(2).Control(9)=   "Label18"
      Tab(2).Control(10)=   "Label51"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Gastos Personales"
      TabPicture(3)   =   "FRolPag.frx":01BE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "Txt3Edad"
      Tab(3).Control(2)=   "TxtDiscap"
      Tab(3).Control(3)=   "TxtVestimenta"
      Tab(3).Control(4)=   "TxtAlimentacion"
      Tab(3).Control(5)=   "TxtEducacion"
      Tab(3).Control(6)=   "TxtSalud"
      Tab(3).Control(7)=   "TxtVivienda"
      Tab(3).Control(8)=   "Label58"
      Tab(3).Control(9)=   "Label60"
      Tab(3).Control(10)=   "Label61"
      Tab(3).Control(11)=   "Label62"
      Tab(3).Control(12)=   "Label63"
      Tab(3).Control(13)=   "Label64"
      Tab(3).Control(14)=   "Label59"
      Tab(3).ControlCount=   15
      Begin VB.CheckBox CheqDecimos 
         Caption         =   "PAGAR DECIMOS EN ROL DE PAGOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7245
         TabIndex        =   138
         Top             =   2940
         Width           =   2955
      End
      Begin MSMask.MaskEdBox MBCta_Diferencia 
         Height          =   330
         Left            =   9240
         TabIndex        =   70
         Top             =   1785
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Fondo_Reserva_G 
         Height          =   330
         Left            =   9240
         TabIndex        =   64
         Top             =   1470
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Fondo_Reserva_P 
         Height          =   330
         Left            =   9240
         TabIndex        =   58
         Top             =   1155
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.CheckBox CheqFondoReserva 
         Caption         =   "PAGAR FONDO DE RESERVA EN ROL DE PAGOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         TabIndex        =   77
         Top             =   2940
         Width           =   2955
      End
      Begin VB.CheckBox CheqRFR 
         Caption         =   "REINGRESO DE FONDOS DE RESERVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   75
         Top             =   2940
         Width           =   2640
      End
      Begin VB.Frame Frame2 
         Caption         =   "DATOS EXTRAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Left            =   -68175
         TabIndex        =   127
         Top             =   420
         Width           =   3795
         Begin VB.TextBox TxtCSSP 
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
            MaxLength       =   15
            TabIndex        =   133
            Top             =   1155
            Width           =   2010
         End
         Begin VB.TextBox TxtAFPONP 
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
            MaxLength       =   10
            TabIndex        =   135
            Top             =   1575
            Width           =   2010
         End
         Begin VB.TextBox TxtNoSeguro 
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
            MaxLength       =   15
            TabIndex        =   129
            Top             =   315
            Width           =   2010
         End
         Begin VB.TextBox TxtCUSSP 
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
            MaxLength       =   12
            TabIndex        =   131
            Top             =   735
            Width           =   2010
         End
         Begin VB.Label Label29 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " No. C.S.S.P."
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
            TabIndex        =   132
            Top             =   1155
            Width           =   1590
         End
         Begin VB.Label Label28 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " A.F.P./ O.N.P."
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
            TabIndex        =   134
            Top             =   1575
            Width           =   1590
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Carnet Es Salud"
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
            TabIndex        =   128
            Top             =   315
            Width           =   1590
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " C.U.S.S.P."
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
            TabIndex        =   130
            Top             =   735
            Width           =   1590
         End
      End
      Begin VB.OptionButton OpcEfectivo 
         Caption         =   "Efectivo"
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
         Left            =   -74790
         TabIndex        =   102
         Top             =   420
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton OpcCheque 
         Caption         =   "Cheque"
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
         Left            =   -71535
         TabIndex        =   104
         Top             =   420
         Width           =   1065
      End
      Begin VB.OptionButton OpcTransferencia 
         Caption         =   "Transferencia"
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
         Left            =   -73425
         TabIndex        =   103
         Top             =   420
         Width           =   1590
      End
      Begin VB.OptionButton OpcOtro 
         Caption         =   "Otros"
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
         Left            =   -70275
         TabIndex        =   105
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox TxtCI 
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
         Left            =   -71955
         MaxLength       =   10
         TabIndex        =   108
         ToolTipText     =   "C01: Cta Cte, C02: Ahorro, C03: Virtual y CI: Acreditar Sueldo a otro Empleado; espacio y la cuenta"
         Top             =   1155
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtCtaAbono 
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
         Left            =   -72795
         MaxLength       =   20
         TabIndex        =   110
         ToolTipText     =   "C01: Cta Cte, C02: Ahorro, C03: Virtual y CI: Acreditar Sueldo a otro Empleado; espacio y la cuenta"
         Top             =   1575
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   126
         Text            =   "0.00"
         Top             =   2940
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   124
         Text            =   "0.00"
         Top             =   2520
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   122
         Text            =   "0.00"
         Top             =   2100
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   120
         Text            =   "0.00"
         Top             =   1680
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   118
         Text            =   "0.00"
         Top             =   1260
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
         Left            =   -73110
         MaxLength       =   10
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Text            =   "0.00"
         Top             =   840
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
         Left            =   -73110
         MaxLength       =   10
         TabIndex        =   114
         Text            =   "0.00"
         Top             =   420
         Width           =   1800
      End
      Begin VB.TextBox TxtDiasDec4to 
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
         Left            =   -64710
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   92
         Text            =   "FRolPag.frx":01DA
         Top             =   1260
         Width           =   540
      End
      Begin VB.TextBox TxtCISustituye 
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
         Left            =   -68070
         MaxLength       =   15
         TabIndex        =   100
         Text            =   "999"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MBCta_Vacaciones_P 
         Height          =   330
         Left            =   5670
         TabIndex        =   56
         Top             =   1155
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Vacaciones_G 
         Height          =   330
         Left            =   5670
         TabIndex        =   62
         Top             =   1470
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox TxtPorcDiscap 
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
         Left            =   -64710
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   98
         Text            =   "FRolPag.frx":01E0
         Top             =   2100
         Width           =   540
      End
      Begin VB.TextBox TxtDiasDec3ro 
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
         Left            =   -70065
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   88
         Text            =   "FRolPag.frx":01E4
         Top             =   1260
         Width           =   540
      End
      Begin MSMask.MaskEdBox MBCta_Aporte_Patronal_G 
         Height          =   330
         Left            =   2100
         TabIndex        =   60
         Top             =   1470
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Quincena 
         Height          =   330
         Left            =   5670
         TabIndex        =   50
         Top             =   735
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Decimo_Cuarto_P 
         Height          =   330
         Left            =   5670
         TabIndex        =   74
         Top             =   2100
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Decimo_Cuarto_G 
         Height          =   330
         Left            =   5670
         TabIndex        =   68
         Top             =   1785
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Antig 
         Height          =   330
         Left            =   9240
         TabIndex        =   52
         Top             =   735
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_IESS_Personal 
         Height          =   330
         Left            =   2100
         TabIndex        =   48
         Top             =   735
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Vacacion 
         Height          =   330
         Left            =   5670
         TabIndex        =   44
         Top             =   420
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Decimo_Tercer_P 
         Height          =   330
         Left            =   2100
         TabIndex        =   72
         Top             =   2100
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Decimo_Tercer_G 
         Height          =   330
         Left            =   2100
         TabIndex        =   66
         Top             =   1785
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_Horas_Ext 
         Height          =   330
         Left            =   9240
         TabIndex        =   46
         Top             =   420
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBCta_IESS_Patronal 
         Height          =   330
         Left            =   2100
         TabIndex        =   54
         Top             =   1155
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox TxtValorDec3ro 
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
         Left            =   -73215
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   86
         Text            =   "FRolPag.frx":01EA
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox TxtValorDec4to 
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
         Left            =   -67860
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   90
         Text            =   "FRolPag.frx":01EE
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox TxtCodProfesion 
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
         Left            =   -71955
         MaxLength       =   15
         TabIndex        =   82
         Text            =   "000000000000000"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxtFPDec 
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
         Left            =   -67755
         MaxLength       =   1
         TabIndex        =   84
         Text            =   "A"
         ToolTipText     =   "Si pago Directo: P, Si pago por Transferencia: A, Si pago en MRL: D"
         Top             =   840
         Width           =   330
      End
      Begin MSMask.MaskEdBox MBCta_Sueldo 
         Height          =   330
         Left            =   2100
         TabIndex        =   42
         ToolTipText     =   "<Ctrl+R> Seleccionar Cuentas de otros Grupo del Rol, <Ctrl+T> Seleccionar Cuentas de otros Grupo a todos los del Rol"
         Top             =   420
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MBFechaVF 
         Height          =   330
         Left            =   -71850
         TabIndex        =   80
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
      Begin MSMask.MaskEdBox MBFechaVI 
         Height          =   330
         Left            =   -73110
         TabIndex        =   79
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
      Begin MSDataListLib.DataCombo DCAplicaConvenio 
         Bindings        =   "FRolPag.frx":01F2
         DataSource      =   "AdoAplicaConvenio"
         Height          =   345
         Left            =   -69645
         TabIndex        =   96
         Top             =   2100
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   609
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCCondiciones 
         Bindings        =   "FRolPag.frx":0212
         DataSource      =   "AdoCondiciones"
         Height          =   345
         Left            =   -69645
         TabIndex        =   94
         Top             =   1680
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   609
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCFormaPago 
         Bindings        =   "FRolPag.frx":022F
         DataSource      =   "AdoFormaPago"
         Height          =   345
         Left            =   -74790
         TabIndex        =   106
         Top             =   735
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   609
         _Version        =   393216
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCTipoBanco 
         Bindings        =   "FRolPag.frx":024A
         DataSource      =   "AdoTipoBanco"
         Height          =   345
         Left            =   -74790
         TabIndex        =   112
         Top             =   2310
         Visible         =   0   'False
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   609
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MBCta_ExtConyugue 
         Height          =   330
         Left            =   9240
         TabIndex        =   139
         Top             =   2100
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo DCSubModulo 
         Bindings        =   "FRolPag.frx":0265
         DataSource      =   "AdoSubCtas"
         Height          =   315
         Left            =   3675
         TabIndex        =   141
         Top             =   2520
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "NOMINA SIN SUBMODULO"
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
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ext. de Conyugue (P)"
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
         TabIndex        =   140
         Top             =   2100
         Width           =   2010
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hora no Trabajadas"
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
         TabIndex        =   69
         Top             =   1785
         Width           =   2010
      End
      Begin VB.Label Label37 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fondo de Reserva (G)"
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
         TabIndex        =   63
         Top             =   1470
         Width           =   2010
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fondo de Reserva (P)"
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
         TabIndex        =   57
         Top             =   1155
         Width           =   2010
      End
      Begin VB.Label Label50 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Acreditar CI de Otro Empleado"
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
         Left            =   -74790
         TabIndex        =   107
         Top             =   1155
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cta. de Transferencia"
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
         Left            =   -74790
         TabIndex        =   109
         Top             =   1575
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label51 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cuenta del Banco a Acreditar:"
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
         Left            =   -74790
         TabIndex        =   111
         Top             =   1995
         Visible         =   0   'False
         Width           =   9150
      End
      Begin VB.Label Label58 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Tercera Edad  "
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
         Left            =   -74895
         TabIndex        =   125
         Top             =   2940
         Width           =   1800
      End
      Begin VB.Label Label60 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Vestimenta"
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
         Left            =   -74895
         TabIndex        =   121
         Top             =   2100
         Width           =   1800
      End
      Begin VB.Label Label61 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Alimentación"
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
         Left            =   -74895
         TabIndex        =   119
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label62 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Educación"
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
         Left            =   -74895
         TabIndex        =   117
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Label63 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Salud"
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
         Left            =   -74895
         TabIndex        =   115
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label64 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Vivienda"
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
         Left            =   -74895
         TabIndex        =   113
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label59 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " (-) Discapacitados  "
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
         Left            =   -74895
         TabIndex        =   123
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label46 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Décimo 4to."
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
         Left            =   -69540
         TabIndex        =   89
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label48 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Días Trabajados"
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
         Left            =   -71745
         TabIndex        =   87
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label49 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Días Trabajados"
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
         Left            =   -66390
         TabIndex        =   91
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label LblTDsustituye 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
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
         Left            =   -66390
         TabIndex        =   101
         Top             =   2520
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label52 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Prov. Vacaciones (P)"
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
         TabIndex        =   55
         Top             =   1155
         Width           =   2010
      End
      Begin VB.Label Label57 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Identificacion de la persona con discapacidad a quien sustituye o Representa"
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
         TabIndex        =   99
         Top             =   2520
         Visible         =   0   'False
         Width           =   6840
      End
      Begin VB.Label Label56 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Porc. Discapacidad"
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
         Left            =   -66600
         TabIndex        =   97
         Top             =   2100
         Width           =   1905
      End
      Begin VB.Label Label55 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Aplica convenio para evitar doble Imposicion"
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
         TabIndex        =   95
         Top             =   2100
         Width           =   5265
      End
      Begin VB.Label Label54 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Condiciones del Trabajador con respecto a discapacidades"
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
         TabIndex        =   93
         Top             =   1680
         Width           =   5265
      End
      Begin VB.Label Label53 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Prov. Vacaciones (G)"
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
         TabIndex        =   61
         Top             =   1470
         Width           =   2010
      End
      Begin VB.Label Label42 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sub-Cuenta"
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
         Top             =   2520
         Width           =   3585
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Aporte Patronal (G)"
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
         TabIndex        =   59
         Top             =   1470
         Width           =   2010
      End
      Begin VB.Label Label41 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Quincena (A-CxC)"
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
         TabIndex        =   49
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Decimo Cuarto (P)"
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
         TabIndex        =   73
         Top             =   2100
         Width           =   2010
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Decimo Cuarto (G)"
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
         TabIndex        =   67
         Top             =   1785
         Width           =   2010
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Antigüedad"
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
         TabIndex        =   51
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Aporte Personal (P)"
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
         TabIndex        =   47
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sueldo Vacacion (G)"
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
         TabIndex        =   43
         Top             =   420
         Width           =   2010
      End
      Begin VB.Label Label39 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Decimo Tercer (P)"
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
         Top             =   2100
         Width           =   2010
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Decimo Tercer (G)"
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
         TabIndex        =   65
         Top             =   1785
         Width           =   2010
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Horas Extras (G)"
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
         TabIndex        =   45
         Top             =   420
         Width           =   2010
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Aporte Patronal (P)"
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
         TabIndex        =   53
         Top             =   1155
         Width           =   2010
      End
      Begin VB.Label Label47 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Décimo 3ro."
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
         TabIndex        =   85
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cod. Profesión (Tabla sectoral)"
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
         TabIndex        =   81
         Top             =   840
         Width           =   2955
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Forma Pago de los décimos"
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
         Left            =   -70275
         TabIndex        =   83
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Vacacion"
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
         TabIndex        =   78
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sueldo Normal (G) "
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
         TabIndex        =   41
         Top             =   420
         Width           =   2010
      End
   End
   Begin VB.TextBox TxtApPer 
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
      Left            =   9345
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "FRolPag.frx":027E
      ToolTipText     =   "<Ctrl+P> Cambia el Aporte Patronal"
      Top             =   2100
      Width           =   750
   End
   Begin VB.TextBox TxtPorcAp 
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
      Left            =   9345
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "FRolPag.frx":0282
      ToolTipText     =   "<Ctrl+P> Cambia el Aporte Personal"
      Top             =   1680
      Width           =   750
   End
   Begin VB.TextBox TxtTarjeta 
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
      IMEMode         =   3  'DISABLE
      Left            =   5985
      PasswordChar    =   "*"
      TabIndex        =   18
      Top             =   1260
      Width           =   1380
   End
   Begin VB.TextBox TxtUsuario 
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
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1260
      Width           =   1065
   End
   Begin VB.TextBox TxtHoras 
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
      Left            =   1995
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "FRolPag.frx":0286
      Top             =   945
      Width           =   1065
   End
   Begin VB.TextBox TxtValorH 
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
      Left            =   4515
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "FRolPag.frx":028A
      Top             =   945
      Width           =   1380
   End
   Begin VB.TextBox TxtIngLiq 
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
      Left            =   7455
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "FRolPag.frx":028E
      Top             =   525
      Width           =   2640
   End
   Begin VB.OptionButton OpcSi 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Si"
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
      TabIndex        =   3
      Top             =   525
      Width           =   540
   End
   Begin VB.OptionButton OpcNo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "No"
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
      TabIndex        =   4
      Top             =   525
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Aceptar"
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
      Left            =   10185
      Picture         =   "FRolPag.frx":0292
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancelar"
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
      Left            =   10185
      Picture         =   "FRolPag.frx":0B5C
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   1050
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   4515
      TabIndex        =   6
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   525
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
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   210
      Top             =   3150
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
      Caption         =   "CxCxP"
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
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   330
      Left            =   3990
      Top             =   3465
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
      Caption         =   "Mes"
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
      Top             =   3465
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
   Begin MSAdodcLib.Adodc AdoRubros 
      Height          =   330
      Left            =   3990
      Top             =   3150
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
      Caption         =   "Rubros"
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
   Begin MSAdodcLib.Adodc AdoGrupos 
      Height          =   330
      Left            =   2100
      Top             =   3150
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
      Caption         =   "Grupos"
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
   Begin MSAdodcLib.Adodc AdoSubCtas 
      Height          =   330
      Left            =   2100
      Top             =   3465
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
      Caption         =   "SubCtas"
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
   Begin MSAdodcLib.Adodc AdoFormaPago 
      Height          =   330
      Left            =   5880
      Top             =   3150
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
      Caption         =   "FormaPago"
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
   Begin MSAdodcLib.Adodc AdoTipoBanco 
      Height          =   330
      Left            =   5880
      Top             =   3465
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
      Caption         =   "TipoBanco"
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
   Begin MSAdodcLib.Adodc AdoAplicaConvenio 
      Height          =   330
      Left            =   7770
      Top             =   3150
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
      Caption         =   "AplicaConvenio"
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
   Begin MSAdodcLib.Adodc AdoCondiciones 
      Height          =   330
      Left            =   7770
      Top             =   3465
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
      Caption         =   "Condiciones"
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
      Top             =   3780
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
   Begin VB.Label Label31 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comisión para Facturacion %"
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
      TabIndex        =   23
      Top             =   1575
      Width           =   2850
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sección del Plantel Educativo"
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
      TabIndex        =   25
      Top             =   1995
      Width           =   3585
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Clave de Registro"
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
      Top             =   1575
      Width           =   1905
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombre Corto"
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
      Top             =   1260
      Width           =   1905
   End
   Begin VB.Label LblMes 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Mes Vacacion"
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
      TabIndex        =   20
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vacacion"
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
      TabIndex        =   19
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte Patronal %"
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
      TabIndex        =   36
      Top             =   2100
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aporte Personal %"
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
      TabIndex        =   34
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pase la Tarjeta sobre la Ranura"
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
      TabIndex        =   17
      Top             =   1260
      Width           =   2850
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Grupo del Rol"
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
      Left            =   5985
      TabIndex        =   13
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Horas por Semana"
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
      TabIndex        =   9
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7455
      TabIndex        =   1
      Top             =   105
      Width           =   2640
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Ingreso"
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
      TabIndex        =   5
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Valor por Hora"
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
      TabIndex        =   11
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ingreso Liquido"
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
      TabIndex        =   7
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Salario Neto?"
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
      TabIndex        =   2
      Top             =   525
      Width           =   1485
   End
   Begin VB.Label LblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASIGNACION DE ROL DE PAGOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7260
   End
End
Attribute VB_Name = "FRolPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OpcAplica As String
Dim OpcCondicion As String
Dim OpcCopiarTodo As Boolean

Private Sub CheqExtC_Click()
   If CheqExtC.Value = 1 Then TxtExtC = Format(IESS_Ext * 100, "#,##0.00") Else TxtExtC = "0.00"
End Sub

Private Sub CheqMaternidad_Click()
 If CheqMaternidad.Value <> 0 Then MBFechaM.Visible = True Else MBFechaM.Visible = False
End Sub

Private Sub CheqSalida_Click()
 If CheqSalida.Value <> 0 Then MBFechaC.Visible = True Else MBFechaC.Visible = False
End Sub

Private Sub Command1_Click()
  FechaValida MBFecha
  FechaValida MBFechaC
  FechaValida MBFechaM
  FechaValida MBFechaVI
  FechaValida MBFechaVF
  TextoValido TxtTarjeta
  TextoValido TxtCI, , True
  TextoValido TxtCSSP, True
  TextoValido TxtCUSSP, True
  TextoValido TxtAFPONP, True
  TextoValido TxtApPer, True
  TextoValido TxtIngLiq, True
  TextoValido TxtPorcAp, True
  TextoValido TxtExtC, True
  TextoValido TxtClave, , True
  TextoValido TxtPorcCom, True
  TextoValido TxtUsuario, , True
  TextoValido TxtNoSeguro, , True
  CodigoCli = LblCodigo.Caption
  
  If Val(CodigoPaisEmpleado) = 593 Then
     TxtCISustituye = "999"
     LblTDsustituye = "N"
  End If
  
  NoMeses = Month(MBFechaVI)
  LblMes.Caption = Format(NoMeses, "00") & " " & UCaseStrg(MesesLetras(NoMeses))
  If AdoCxCxP.Recordset.RecordCount > 0 Then
     AdoCxCxP.Recordset.MoveFirst
     AdoCxCxP.Recordset.Find ("Codigo = '" & CodigoCli & "' ")
     If AdoCxCxP.Recordset.EOF Then SetAddNew AdoCxCxP
  Else
     SetAddNew AdoCxCxP
  End If
  
 'Buscamos el codigo del submodulo de gastos
  CodigoInv = ""
  If DCSubModulo.Text = "" Then DCSubModulo.Text = "NOMINA SIN SUBMODULO"
  If AdoSubCtas.Recordset.RecordCount > 0 Then
     AdoSubCtas.Recordset.MoveFirst
     AdoSubCtas.Recordset.Find ("Detalle = '" & DCSubModulo.Text & "' ")
     If Not AdoSubCtas.Recordset.EOF Then CodigoInv = AdoSubCtas.Recordset.fields("Codigo")
  End If
  If CodigoInv = "" Then CodigoInv = Ninguno
  
 'MsgBox CodigoInv & vbCrLf & TxtTarjeta
  SetFields AdoCxCxP, "SubModulo", CodigoInv
  SetFields AdoCxCxP, "Horas_Ext", adFalse
  SetFields AdoCxCxP, "T", Normal
  SetFields AdoCxCxP, "Fecha", MBFecha
  SetFields AdoCxCxP, "Item", NumEmpresa
  SetFields AdoCxCxP, "Codigo", CodigoCli
  SetFields AdoCxCxP, "Valor_Hora", Val(CDbl(TxtValorH))
  SetFields AdoCxCxP, "Horas_Sem", Val(CCur(TxtHoras))
  SetFields AdoCxCxP, "Salario", Val(CCur(TxtIngLiq))
  SetFields AdoCxCxP, "Valor_Dec_3ro", Val(CCur(TxtValorDec3ro))
  SetFields AdoCxCxP, "Valor_Dec_4to", Val(CCur(TxtValorDec4to))
  SetFields AdoCxCxP, "Dias_Dec_3ro", Val(TxtDiasDec3ro)
  SetFields AdoCxCxP, "Dias_Dec_4to", Val(TxtDiasDec4to)
  SetFields AdoCxCxP, "Porc_Com", Val(CSng(TxtPorcCom) / 100)
  SetFields AdoCxCxP, "Usuario", TxtUsuario
  SetFields AdoCxCxP, "Clave", TxtClave
  SetFields AdoCxCxP, "Grupo_Rol", DCGrupos
  SetFields AdoCxCxP, "Tarjeta", Sin_Signos_Especiales(TxtTarjeta)
  If OpcSi.Value Then SetFields AdoCxCxP, "SN", "2" Else SetFields AdoCxCxP, "SN", "1"
  Cta = CambioCodigoCta(MBCta_Horas_Ext)
  If Val(MidStrg(Cta, 1, 1)) > 0 Then SetFields AdoCxCxP, "Horas_Ext", adTrue
  SetFields AdoCxCxP, "No_CSSP", UCaseStrg(TxtCSSP)
  SetFields AdoCxCxP, "No_CUSSP", UCaseStrg(TxtCUSSP)
  SetFields AdoCxCxP, "AFP_ONP", UCaseStrg(TxtAFPONP)
  SetFields AdoCxCxP, "No_Personal", TxtNoSeguro
  
'''  SetFields AdoCxCxP, "Cta_Sueldo", CambioCodigoCta(MBCta_Sueldo)
'''  SetFields AdoCxCxP, "Cta_Horas_Ext", CambioCodigoCta(MBCta_Horas_Ext)
'''  SetFields AdoCxCxP, "Cta_Antiguedad", CambioCodigoCta(MBCta_Antig)
'''  SetFields AdoCxCxP, "Cta_Diferencia", CambioCodigoCta(MBCta_Diferencia)
'''  SetFields AdoCxCxP, "Cta_Vacacion", CambioCodigoCta(MBCta_Vacacion)
'''  SetFields AdoCxCxP, "Cta_Aporte_Patronal_G", CambioCodigoCta(MBCta_Aporte_Patronal_G)
'''  SetFields AdoCxCxP, "Cta_Decimo_Cuarto_G", CambioCodigoCta(MBCta_Decimo_Cuarto_G)
'''  SetFields AdoCxCxP, "Cta_Decimo_Cuarto_P", CambioCodigoCta(MBCta_Decimo_Cuarto_P)
'''  SetFields AdoCxCxP, "Cta_Decimo_Tercer_G", CambioCodigoCta(MBCta_Decimo_Tercer_G)
'''  SetFields AdoCxCxP, "Cta_Decimo_Tercer_P", CambioCodigoCta(MBCta_Decimo_Tercer_P)
'''  SetFields AdoCxCxP, "Cta_Fondo_Reserva_G", CambioCodigoCta(MBCta_Fondo_Reserva_G)
'''  SetFields AdoCxCxP, "Cta_Fondo_Reserva_P", CambioCodigoCta(MBCta_Fondo_Reserva_P)
'''  SetFields AdoCxCxP, "Cta_Ext_Conyugue_P", CambioCodigoCta(MBCta_ExtConyugue)
'''  SetFields AdoCxCxP, "Cta_Vacaciones_G", CambioCodigoCta(MBCta_Vacaciones_G)
'''  SetFields AdoCxCxP, "Cta_Vacaciones_P", CambioCodigoCta(MBCta_Vacaciones_P)
'''  SetFields AdoCxCxP, "Cta_IESS_Patronal", CambioCodigoCta(MBCta_IESS_Patronal)
'''  SetFields AdoCxCxP, "Cta_IESS_Personal", CambioCodigoCta(MBCta_IESS_Personal)
'''  SetFields AdoCxCxP, "Cta_Quincena", CambioCodigoCta(MBCta_Quincena)
  SetFields AdoCxCxP, "CodProfesion", Format$(Val(TxtCodProfesion), "0000000000")
  SetFields AdoCxCxP, "FormaPago10to", UCaseStrg(TxtFPDec)
  SetFields AdoCxCxP, "Ejecutivo", NombreCliente
  
  SetFields AdoCxCxP, "FechaVI", MBFechaVI
  SetFields AdoCxCxP, "FechaVF", MBFechaVF
  SetFields AdoCxCxP, "Mes", Val(SinEspaciosIzq(LblMes.Caption))
  SetFields AdoCxCxP, "Cta_Transferencia", Ninguno
  SetFields AdoCxCxP, "Acreditar_Cta", Ninguno
  SetFields AdoCxCxP, "TC", TipoCta
  SetFields AdoCxCxP, "Cta_Forma_Pago", TrimStrg(SinEspaciosIzq(DCFormaPago))
  SetFields AdoCxCxP, "Pagar_Fondo_Reserva", CBool(CheqFondoReserva.Value)
  SetFields AdoCxCxP, "Pagar_Decimos", CBool(CheqDecimos.Value)
  SetFields AdoCxCxP, "TiempoParcial", CBool(CheqTiempoParcial.Value)
  SetFields AdoCxCxP, "Reingreso_FR", CBool(CheqRFR.Value)
  SetFields AdoCxCxP, "ExtC", CBool(CheqExtC.Value)
  SetFields AdoCxCxP, "Identificacion", TxtCISustituye
  SetFields AdoCxCxP, "TIdentificacion", LblTDsustituye.Caption
  SetFields AdoCxCxP, "Aplica", OpcAplica
  SetFields AdoCxCxP, "Condicion", OpcCondicion
  SetFields AdoCxCxP, "Porcentaje", Val(CCur(TxtPorcDiscap))
  SetFields AdoCxCxP, "Vivienda", Val(CCur(TxtVivienda))
  SetFields AdoCxCxP, "Salud", Val(CCur(TxtSalud))
  SetFields AdoCxCxP, "Educacion", Val(CCur(TxtEducacion))
  SetFields AdoCxCxP, "Alimentacion", Val(CCur(TxtAlimentacion))
  SetFields AdoCxCxP, "Vestimenta", Val(CCur(TxtVestimenta))
  SetFields AdoCxCxP, "Discapacidad", Val(CCur(TxtDiscap))
  SetFields AdoCxCxP, "Tercera_Edad", Val(CCur(Txt3Edad))
  SetFields AdoCxCxP, "Cod_Ejec", Abreviatura_Texto(LblCliente.Caption)
  SetFields AdoCxCxP, "Item", NumEmpresa
  SetFields AdoCxCxP, "Periodo", Periodo_Contable
  
  If CheqSalida.Value <> 0 Then
     SetFields AdoCxCxP, "T", "R"
     SetFields AdoCxCxP, "FechaC", MBFechaC
  Else
     SetFields AdoCxCxP, "T", Normal
     SetFields AdoCxCxP, "FechaC", FechaSistema
  End If
  
  If CheqMaternidad.Value <> 0 Then
     SetFields AdoCxCxP, "FechaMat", MBFechaM
  Else
     SetFields AdoCxCxP, "FechaMat", FechaSistema
  End If
  SetFields AdoCxCxP, "Maternidad", CBool(CheqMaternidad.Value)
  
  If OpcEfectivo.Value Then
     SetFields AdoCxCxP, "FP", "E"
  ElseIf OpcCheque.Value Then
     SetFields AdoCxCxP, "FP", "C"
  ElseIf OpcTransferencia.Value Then
     SetFields AdoCxCxP, "FP", "T"
     SetFields AdoCxCxP, "Cta_Transferencia", TxtCtaAbono
     SetFields AdoCxCxP, "Acreditar_Cta", TxtCI
     If AdoTipoBanco.Recordset.RecordCount > 0 Then
        AdoTipoBanco.Recordset.MoveFirst
        AdoTipoBanco.Recordset.Find ("Descripcion = '" & DCTipoBanco.Text & "' ")
        If Not AdoTipoBanco.Recordset.EOF Then
           SetFields AdoCxCxP, "Codigo_Banco", AdoTipoBanco.Recordset.fields("Codigo")
        End If
     End If
  Else
     SetFields AdoCxCxP, "FP", "O"
  End If
  SetUpdate AdoCxCxP
 'Crear Clave de Ingreso al Sistema
  sSQL = "SELECT TODOS, Codigo, Nombre_Completo, Usuario,  Clave, Primaria, Secundaria, Bachillerato, Cod_Ejec, ID " _
       & "FROM Accesos " _
       & "WHERE Codigo = '" & CodigoCli & "' "
  Select_Adodc AdoCxCxP, sSQL
  With AdoCxCxP.Recordset
   If .RecordCount > 0 Then
      .fields("Clave") = TxtClave
      .fields("Usuario") = TxtUsuario
      .fields("Nombre_Completo") = ULCase(NombreCliente)
      .fields("Cod_Ejec") = Abreviatura_Texto(NombreCliente)
      .Update
   Else
      SetAdoAddNew "Accesos", True
      SetAdoFields "TODOS", True
      SetAdoFields "Clave", TxtClave
      SetAdoFields "Codigo", CodigoCli
      SetAdoFields "Usuario", TxtUsuario
      SetAdoFields "Nombre_Completo", ULCase(NombreCliente)
      SetAdoFields "Cod_Ejec", Abreviatura_Texto(NombreCliente)
      If CheqP.Value Then SetAdoFields "Primaria", adTrue
      If CheqS.Value Then SetAdoFields "Secundaria", adTrue
      If CheqB.Value Then SetAdoFields "Bachillerato", adTrue
      SetAdoUpdate
   End If
  End With
  Ejecutar_SQL_SP sSQL
  Unload FRolPago
End Sub

Private Sub Command2_Click()
  Unload FRolPago
End Sub

Private Sub DCAplicaConvenio_LostFocus()
  OpcAplica = "NA"
  With AdoAplicaConvenio.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & DCAplicaConvenio & "' ")
       If Not .EOF Then OpcAplica = .fields("Codigo")
   End If
  End With
End Sub

Private Sub DCCondiciones_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCondiciones_LostFocus()
Dim EsVisible As Boolean
  OpcCondicion = "01"
  With AdoCondiciones.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & DCCondiciones & "' ")
       If Not .EOF Then OpcCondicion = .fields("Codigo")
   End If
  End With
  If Val(OpcCondicion) > 2 Then EsVisible = True Else EsVisible = False
  Label57.Visible = EsVisible
  TxtCISustituye.Visible = EsVisible
  LblTDsustituye.Visible = EsVisible
End Sub

Private Sub DCFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFormaPago_LostFocus()
  With AdoFormaPago.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cuentas = '" & DCFormaPago & "' ")
       If Not .EOF Then TipoCta = .fields("TC")
   End If
  End With
  If TipoCta = Ninguno Then MsgBox "Vuelva a elejir la cuenta, no ha seleccionado correctamente "
End Sub

Private Sub DCGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupos_LostFocus()
  Leer_Catalogo_Rol_Pagos DCGrupos.Text
End Sub

Private Sub DCSubModulo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipoBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
Dim IdRol As Integer
  Datos_IESS FechaSistema
  
  sSQL = "SELECT Grupo_Rol, Cta_Diferencia, Cta_Vacacion, Cta_Sueldo, Cta_Horas_Ext, Cta_Aporte_Patronal_G, " _
       & "Cta_Decimo_Cuarto_G, Cta_Decimo_Cuarto_P, Cta_Decimo_Tercer_P, Cta_Fondo_Reserva_G, Cta_Fondo_Reserva_P, " _
       & "Cta_IESS_Personal, Cta_Quincena, Cta_Decimo_Tercer_G,Cta_IESS_Patronal, Cta_Antiguedad, " _
       & "Cta_Vacaciones_G, Cta_Vacaciones_P, Cta_Ext_Conyugue_P " _
       & "FROM Catalogo_Rol_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Grupo_Rol "
  SelectDB_Combo DCGrupos, AdoGrupos, sSQL, "Grupo_Rol"

  sSQL = "SELECT * " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'APLICA CONVENIO' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCAplicaConvenio, AdoAplicaConvenio, sSQL, "Descripcion", True
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'TRABAJADOR' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCondiciones, AdoCondiciones, sSQL, "Descripcion"
  
'''  sSQL = "SELECT * " _
'''       & "FROM Catalogo_Rol_Pagos " _
'''       & "WHERE Item = '" & NumEmpresa & "' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc AdoMes, sSQL
'''  With AdoMes.Recordset
'''       For IdRol = 0 To .Fields.Count - 1
'''           If MidStrg(.Fields(IdRol).Name, 1, 3) = "Cta" Then
'''              Select Case .Fields(IdRol).Name
'''                Case "Cta_Forma_Pago"
'''                     SQL2 = "UPDATE Catalogo_Rol_Pagos " _
'''                          & "SET FP = 'E', TC = 'CJ', " _
'''                          & "Cta_Forma_Pago = '" & Cta_General & "', " _
'''                          & "Pagar_Fondo_Reserva = " & Val(adTrue) & " " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND Cta_Forma_Pago = '.' "
'''                Case "Cta_Transferencia"
'''                     SQL2 = "UPDATE Catalogo_Rol_Pagos " _
'''                          & "SET Cta_Transferencia = '.' " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND Cta_Transferencia = ' ' "
'''                Case Else
'''                     SQL2 = "UPDATE Catalogo_Rol_Pagos " _
'''                          & "SET " & .Fields(IdRol).Name & " = '0' " _
'''                          & "WHERE Item = '" & NumEmpresa & "' " _
'''                          & "AND Periodo = '" & Periodo_Contable & "' " _
'''                          & "AND LEN(" & .Fields(IdRol).Name & ") <= 1 " _
'''                          & "AND " & .Fields(IdRol).Name & " <> '0' "
'''              End Select
'''              Ejecutar_SQL_SP SQL2
'''           End If
'''       Next IdRol
'''  End With
  sSQL = "SELECT * " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "AND Descripcion <> '.' " _
       & "ORDER BY Descripcion "
  SelectDB_Combo DCTipoBanco, AdoTipoBanco, sSQL, "Descripcion"
  DCTipoBanco.Text = "NINGUN BANCO ASIGNADO"
  
  LblCodigo.Caption = CodigoCli
  LblCliente.Caption = NombreCliente
  FormatoMaskCta MBCta_Antig

  FormatoMaskCta MBCta_Sueldo
  FormatoMaskCta MBCta_Vacacion
  FormatoMaskCta MBCta_Horas_Ext
  FormatoMaskCta MBCta_Diferencia
  FormatoMaskCta MBCta_Aporte_Patronal_G
  FormatoMaskCta MBCta_Decimo_Cuarto_G
  FormatoMaskCta MBCta_Decimo_Cuarto_P
  FormatoMaskCta MBCta_Decimo_Tercer_G
  FormatoMaskCta MBCta_Decimo_Tercer_P
  FormatoMaskCta MBCta_Fondo_Reserva_G
  FormatoMaskCta MBCta_Fondo_Reserva_P
  FormatoMaskCta MBCta_IESS_Patronal
  FormatoMaskCta MBCta_IESS_Personal
  FormatoMaskCta MBCta_Vacaciones_G
  FormatoMaskCta MBCta_Vacaciones_P
  FormatoMaskCta MBCta_Quincena
  FormatoMaskCta MBCta_ExtConyugue
  
 'Colocamos los abreviados de los rubros del Rol
  MBFecha.Text = FechaSistema
  MBFechaM.Visible = False
  MBFechaC.Visible = False
  CheqP.Value = 0
  CheqS.Value = 0
  CheqB.Value = 0
  CheqExtC.Value = 0
  CheqTiempoParcial.Value = 0
  CheqMaternidad.Value = 0
  CheqSalida.Value = 0
  CheqRFR.Value = 0
  CheqFondoReserva.Value = 0
  CheqDecimos.Value = 0

'  TxtApPer.Text = "0.00"
'  TxtPorcAp.Text = "0.00"
  TxtExtC.Text = "0.00"
  TxtValorH.Text = "0"
  TxtHoras.Text = "0"
  TxtIngLiq.Text = "0"
  TxtCtaAbono.Text = ""

  sSQL = "SELECT * " _
       & "FROM Accesos " _
       & "WHERE Codigo = '" & CodigoCli & "' "
  Select_Adodc AdoSubCtas, sSQL
  With AdoSubCtas.Recordset
   If .RecordCount > 0 Then
       If .fields("Primaria") Then CheqP.Value = 1
       If .fields("Secundaria") Then CheqS.Value = 1
       If .fields("Bachillerato") Then CheqB.Value = 1
       TxtUsuario = .fields("Usuario")
       TxtClave = .fields("Clave")
   End If
  End With
    
  sSQL = "SELECT TC, Codigo, Detalle " _
       & "FROM Catalogo_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Agrupacion = " & adFalse & " " _
       & "AND TC IN ('CC','G','GC') " _
       & "ORDER BY Detalle, TC "
  SelectDB_Combo DCSubModulo, AdoSubCtas, sSQL, "Detalle"
  DCSubModulo.Text = "NOMINA SIN SUBMODULO"
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & CodigoCli & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoCxCxP, sSQL
  With AdoCxCxP.Recordset
   If .RecordCount > 0 Then
        DCGrupos = .fields("Grupo_Rol")
        If .fields("T") = Normal Then
            LblCodigo.BackColor = Negro
            LblCliente.BackColor = Negro
        Else
            LblCodigo.BackColor = Rojo
            LblCliente.BackColor = Rojo
        End If
        MBFecha = .fields("Fecha")
        MBFechaC = .fields("FechaC")
        MBFechaM = .fields("FechaMat")
        MBFechaVI = .fields("FechaVI")
        MBFechaVF = .fields("FechaVF")
        NoMeses = Month(MBFechaVI)
        LblMes.Caption = Format(NoMeses, "00") & " " & UCaseStrg(MesesLetras(NoMeses))
        TxtPorcCom = Format(.fields("Porc_Com") * 100, "#,##0.00")
        TxtValorH = Format(.fields("Valor_Hora"), "#,##0.00000")
        TxtHoras = Format(.fields("Horas_Sem"), "#,##0.00")
        TxtIngLiq = Format(.fields("Salario"), "#,##0.00")
        TxtValorDec3ro = Format(.fields("Valor_Dec_3ro"), "#,##0.00")
        TxtValorDec4to = Format(.fields("Valor_Dec_4to"), "#,##0.00")
        TxtDiasDec3ro = Format(.fields("Dias_Dec_3ro"), "000")
        TxtDiasDec4to = Format(.fields("Dias_Dec_4to"), "000")
        
        TxtNoSeguro = .fields("No_Personal")
        TxtCtaAbono = .fields("Cta_Transferencia")
        TxtCSSP = .fields("No_CSSP")
        TxtCUSSP = .fields("No_CUSSP")
        TxtAFPONP = .fields("AFP_ONP")
        TxtCodProfesion = Format(Val(.fields("CodProfesion")), "0000000000")
        TxtFPDec = UCaseStrg(.fields("FormaPago10to"))
        TxtCI = .fields("Acreditar_Cta")
        
        TxtVivienda = .fields("Vivienda")
        TxtSalud = .fields("Salud")
        TxtEducacion = .fields("Educacion")
        TxtAlimentacion = .fields("Alimentacion")
        TxtVestimenta = .fields("Vestimenta")
        TxtDiscap = .fields("Discapacidad")
        Txt3Edad = .fields("Tercera_Edad")
        
        TxtCISustituye = .fields("Identificacion")
        LblTDsustituye.Caption = .fields("TIdentificacion")
        TxtPorcDiscap = .fields("Porcentaje")
        OpcAplica = .fields("Aplica")
        OpcCondicion = .fields("Condicion")
              
        If TxtFPDec = Ninguno Then TxtFPDec = "A"
        
        With AdoCondiciones.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Codigo = '" & OpcCondicion & "' ")
             If Not .EOF Then DCCondiciones = .fields("Descripcion")
         End If
        End With
        
        With AdoAplicaConvenio.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Codigo = '" & OpcAplica & "' ")
             If Not .EOF Then DCAplicaConvenio = .fields("Descripcion")
         End If
        End With
        
        If .fields("Pagar_Fondo_Reserva") Then CheqFondoReserva.Value = 1
        If .fields("Pagar_Decimos") Then CheqDecimos.Value = 1
        If .fields("Reingreso_FR") Then CheqRFR.Value = 1
        If .fields("TiempoParcial") Then CheqTiempoParcial.Value = 1
        If .fields("ExtC") Then
            TxtExtC = Format(IESS_Ext * 100, "#,##0.00")
            CheqExtC.Value = 1
        End If
        If .fields("Maternidad") Then
            MBFechaM.Visible = True
            CheqMaternidad.Value = 1
        End If
        If .fields("T") = "R" Then
            MBFechaC.Visible = True
            CheqSalida.Value = 1
        End If
                
        TipoCta = .fields("TC")
        Cta_Aux = .fields("Cta_Forma_Pago")
        TxtTarjeta = .fields("Tarjeta")
        If .fields("SN") = "2" Then
            OpcSi.Value = True
            'TxtApPer = Format((IESS_Pat + IESS_Per) * 100, "#,##0.00")
        Else
            OpcNo.Value = True
            'TxtApPer = Format(IESS_Pat * 100, "#,##0.00")
            'TxtPorcAp = Format(IESS_Per * 100, "#,##0.00")
        End If
        If AdoSubCtas.Recordset.RecordCount > 0 Then
           AdoSubCtas.Recordset.MoveFirst
           AdoSubCtas.Recordset.Find ("Codigo = '" & .fields("SubModulo") & "' ")
           If Not AdoSubCtas.Recordset.EOF Then DCSubModulo.Text = AdoSubCtas.Recordset.fields("Detalle")
        End If
        Select Case .fields("FP")
          Case "E": OpcEfectivo.Value = True
          Case "C": OpcCheque.Value = True
          Case "T": OpcTransferencia.Value = True
          Case "O": OpcOtro.Value = True
        End Select
        If .fields("Codigo_Banco") > 0 Then
            If AdoTipoBanco.Recordset.RecordCount > 0 Then
               AdoTipoBanco.Recordset.MoveFirst
               AdoTipoBanco.Recordset.Find ("Codigo = '" & .fields("Codigo_Banco") & "' ")
               If Not AdoTipoBanco.Recordset.EOF Then
                  DCTipoBanco.Text = AdoTipoBanco.Recordset.fields("Descripcion")
               End If
            End If
        End If
   End If
  End With
  
  Leer_Catalogo_Rol_Pagos DCGrupos.Text
  
 'Funcion Tipo de Pago
  Tipo_de_Pago
  With AdoFormaPago.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo = '" & Cta_Aux & "' ")
       If Not .EOF Then DCFormaPago = .fields("Codigo") & " - " & .fields("Cuenta")
   End If
  End With
  If OpcSi.Value Then OpcSi.SetFocus Else OpcNo.SetFocus
  If Val(CodigoPaisEmpleado) <> 593 Then
     Label55.Visible = True
     DCAplicaConvenio.Visible = True
  End If
End Sub

Private Sub Form_Load()
  CentrarForm FRolPago
  ConectarAdodc AdoAux
  ConectarAdodc AdoMes
  ConectarAdodc AdoGrupos
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoRubros
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoSubCtas
  ConectarAdodc AdoTipoBanco
  ConectarAdodc AdoFormaPago
  ConectarAdodc AdoCondiciones
  ConectarAdodc AdoAplicaConvenio
  FCtaAhorro.Caption = "ASIGNACION A ROL DE PAGOS"
End Sub

Private Sub MBCta_Sueldo_KeyDown(KeyCode As Integer, Shift As Integer)
'''  Keys_Especiales Shift
'''  PresionoEnter KeyCode
'''  If CtrlDown And KeyCode = vbKeyR Then
'''
'''     OpcCopiarTodo = False
'''     FrmGrupos.Visible = True
'''     DLGrupos.SetFocus
'''  End If
'''  If CtrlDown And KeyCode = vbKeyT Then
'''     OpcCopiarTodo = True
'''     FrmGrupos.Visible = True
'''     DLGrupos.SetFocus
'''  End If
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub MBFechaC_GotFocus()
  MarcarTexto MBFechaC
End Sub

Private Sub MBFechaC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaC_LostFocus()
  FechaValida MBFechaC
End Sub

Private Sub MBFechaM_GotFocus()
   MarcarTexto MBFechaM
End Sub

Private Sub MBFechaM_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBFechaM_LostFocus()
  FechaValida MBFechaM
End Sub

Private Sub MBFechaVF_GotFocus()
  MarcarTexto MBFechaVF
End Sub

Private Sub MBFechaVF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaVF_LostFocus()
  FechaValida MBFechaVF
End Sub

Private Sub MBFechaVI_GotFocus()
  MarcarTexto MBFechaVI
End Sub

Private Sub MBFechaVI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaVI_LostFocus()
  FechaValida MBFechaVI
  NoMeses = Month(MBFechaVI)
  LblMes.Caption = Format(NoMeses, "00") & " " & UCaseStrg(MesesLetras(NoMeses))
End Sub

Private Sub OpcCheque_Click()
 Tipo_de_Pago
End Sub

Private Sub OpcEfectivo_Click()
  Tipo_de_Pago
End Sub

Private Sub OpcOtro_Click()
  Tipo_de_Pago
End Sub

Private Sub OpcSi_Click()
'    TxtPorcAp = "0.00"
'    TxtApPer = Format((IESS_Pat + IESS_Per) * 100, "#,##0.00")
End Sub

Private Sub OpcNo_Click()
''    TxtApPer = Format(IESS_Pat * 100, "#,##0.00")
''    TxtPorcAp = Format(IESS_Per * 100, "#,##0.00")
End Sub

Private Sub OpcTransferencia_Click()
  Tipo_de_Pago
End Sub

Private Sub Txt3Edad_GotFocus()
  MarcarTexto Txt3Edad
End Sub

Private Sub Txt3Edad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Txt3Edad_LostFocus()
  TextoValido Txt3Edad, True, , 2
End Sub

Private Sub TxtAFPONP_GotFocus()
  MarcarTexto TxtAFPONP
End Sub

Private Sub TxtAFPONP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAlimentacion_GotFocus()
  MarcarTexto TxtAlimentacion
End Sub

Private Sub TxtAlimentacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAlimentacion_LostFocus()
  TextoValido TxtAlimentacion, True, , 2
End Sub

Private Sub TxtApPer_LostFocus()
  TextoValido TxtApPer, True
End Sub

Private Sub TxtCI_GotFocus()
  MarcarTexto TxtCI
End Sub

Private Sub TxtCI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCI_LostFocus()
  TextoValido TxtCI, , True
End Sub

Private Sub TxtCISustituye_GotFocus()
  MarcarTexto TxtCISustituye
End Sub

Private Sub TxtCISustituye_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCISustituye_LostFocus()
  TextoValido TxtCISustituye, , True
  DigVerif = Digito_Verificador(TxtCISustituye)
  LblTDsustituye = Tipo_RUC_CI.Tipo_Beneficiario
End Sub

Private Sub TxtClave_GotFocus()
  MarcarTexto TxtClave
End Sub

Private Sub TxtClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyF9 Then MsgBox "Clave: " & TxtClave
End Sub

Private Sub TxtClave_LostFocus()
 TextoValido TxtClave, , True
 If TxtClave.Text <> Ninguno Then
    If AdoCxCxP.Recordset.RecordCount > 0 Then
       AdoCxCxP.Recordset.MoveFirst
       AdoCxCxP.Recordset.Find ("Clave = '" & TxtClave.Text & "' ")
       If Not AdoCxCxP.Recordset.EOF Then
          If CodigoCli <> AdoCxCxP.Recordset.fields("Codigo") Then
             MsgBox "Clave ya asignado"
             TxtClave.SetFocus
          End If
       End If
    End If
 End If
End Sub

Private Sub TxtCodProfesion_GotFocus()
 MarcarTexto TxtCodProfesion
End Sub

Private Sub TxtCodProfesion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodProfesion_LostFocus()
  TxtCodProfesion = Format(Val(TxtCodProfesion), "0000000000")
End Sub

Private Sub TxtCSSP_GotFocus()
  MarcarTexto TxtCSSP
End Sub

Private Sub TxtCSSP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCtaAbono_GotFocus()
  MarcarTexto TxtCtaAbono
End Sub

Private Sub TxtCtaAbono_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCtaAbono_LostFocus()
  TextoValido TxtCtaAbono, , True
End Sub

Private Sub TxtCUSSP_GotFocus()
  MarcarTexto TxtCUSSP
End Sub

Private Sub TxtCUSSP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDiasDec3ro_GotFocus()
  MarcarTexto TxtDiasDec3ro
End Sub

Private Sub TxtDiasDec3ro_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyI Then Importar_Rol_Pagos "Decimo_III", False
  PresionoEnter KeyCode
End Sub

Private Sub TxtDiasDec4to_GotFocus()
  MarcarTexto TxtDiasDec4to
End Sub

Private Sub TxtDiasDec4to_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyI Then Importar_Rol_Pagos "Decimo_IV", False
  PresionoEnter KeyCode
End Sub

Private Sub TxtDiscap_GotFocus()
  MarcarTexto TxtDiscap
End Sub

Private Sub TxtDiscap_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDiscap_LostFocus()
  TextoValido TxtDiscap, True, , 2
End Sub

Private Sub TxtEducacion_GotFocus()
  MarcarTexto TxtEducacion
End Sub

Private Sub TxtEducacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEducacion_LostFocus()
  TextoValido TxtEducacion, True, , 2
End Sub

Private Sub TxtFPDec_GotFocus()
   MarcarTexto TxtFPDec
End Sub

Private Sub TxtFPDec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFPDec_LostFocus()
  TextoValido TxtFPDec, False, True
End Sub

Private Sub TxtHoras_GotFocus()
  MarcarTexto TxtHoras
End Sub

Private Sub TxtHoras_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtHoras_LostFocus()
  TextoValido TxtHoras, True
  If Val(TxtHoras.Text) <= 0 Then TxtHoras.Text = "1.00"
  TxtValorH.Text = Format(Val(CCur(TxtIngLiq.Text)) / (Val(CCur(TxtHoras.Text)) * 4), "##0.00000")
End Sub

Private Sub TxtIngLiq_GotFocus()
  MarcarTexto TxtIngLiq
End Sub

Private Sub TxtIngLiq_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtIngLiq_LostFocus()
  TextoValido TxtIngLiq, True
End Sub

Private Sub TxtNoSeguro_GotFocus()
   MarcarTexto TxtNoSeguro
End Sub

Private Sub TxtNoSeguro_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorcCom_GotFocus()
  MarcarTexto TxtPorcCom
End Sub

Private Sub TxtPorcCom_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorcCom_LostFocus()
  TextoValido TxtPorcCom, True
End Sub

Private Sub TxtPorcDiscap_GotFocus()
  MarcarTexto TxtPorcDiscap
End Sub

Private Sub TxtPorcDiscap_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorcDiscap_LostFocus()
  If Val(TxtPorcDiscap) > 100 Then TxtPorcDiscap = "100"
End Sub

Private Sub TxtSalud_GotFocus()
  MarcarTexto TxtSalud
End Sub

Private Sub TxtSalud_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSalud_LostFocus()
  TextoValido TxtSalud, True, , 2
End Sub

Private Sub TxtTarjeta_GotFocus()
  MarcarTexto TxtTarjeta
End Sub

Private Sub TxtTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
'  PresionoEnter KeyCode
End Sub

Private Sub TxtUsuario_GotFocus()
  MarcarTexto TxtUsuario
End Sub

Private Sub TxtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtUsuario_LostFocus()
 TextoValido TxtUsuario, , True
 If TxtUsuario.Text <> Ninguno Then
    If AdoCxCxP.Recordset.RecordCount > 0 Then
       AdoCxCxP.Recordset.MoveFirst
       AdoCxCxP.Recordset.Find ("Usuario = '" & TxtUsuario.Text & "' ")
       If Not AdoCxCxP.Recordset.EOF Then
          If CodigoCli <> AdoCxCxP.Recordset.fields("Codigo") Then
             MsgBox "Usuario ya asignado"
             TxtUsuario.SetFocus
          End If
       End If
    End If
 End If
End Sub

Private Sub TxtValorDec3ro_GotFocus()
  MarcarTexto TxtValorDec3ro
End Sub

Private Sub TxtValorDec3ro_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorDec3ro_LostFocus()
  TextoValido TxtValorDec3ro, True
End Sub

Private Sub TxtValorDec4to_GotFocus()
  MarcarTexto TxtValorDec4to
End Sub

Private Sub TxtValorDec4to_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorDec4to_LostFocus()
  TextoValido TxtValorDec4to, True
End Sub

Private Sub TxtValorH_GotFocus()
  MarcarTexto TxtValorH
End Sub

Private Sub TxtValorH_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtValorH_LostFocus()
  TextoValido TxtValorH
End Sub

Public Sub Tipo_de_Pago()
  DCTipoBanco.Visible = False
  Label51.Visible = False
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas, TC,Codigo,Cuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo < '3' " _
       & "AND DG = 'D' "
  If OpcEfectivo.Value Then
     sSQL = sSQL & "AND TC = 'CJ' "
     Label18.Visible = False
     Label50.Visible = False
     TxtCI.Visible = False
     TxtCtaAbono.Visible = False
  ElseIf OpcCheque.Value Then
     sSQL = sSQL & "AND TC = 'BA' "
     Label18.Visible = False
     Label50.Visible = False
     TxtCI.Visible = False
     TxtCtaAbono.Visible = False
  ElseIf OpcTransferencia.Value Then
     sSQL = sSQL & "AND TC = 'BA' "
     Label18.Visible = True
     TxtCI.Visible = True
     Label50.Visible = True
     TxtCtaAbono.Visible = True
     DCTipoBanco.Visible = True
     Label51.Visible = True
  Else
     sSQL = sSQL & "AND TC NOT IN ('CJ','BA','RI','RF') "
     Label18.Visible = False
     Label50.Visible = False
     TxtCI.Visible = False
     TxtCtaAbono.Visible = False
  End If
  sSQL = sSQL & "ORDER BY Codigo "
  SelectDB_Combo DCFormaPago, AdoFormaPago, sSQL, "Cuentas"
End Sub

Private Sub TxtVestimenta_GotFocus()
   MarcarTexto TxtVestimenta
End Sub

Private Sub TxtVestimenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtVestimenta_LostFocus()
  TextoValido TxtVestimenta, True, , 2
End Sub

Private Sub TxtVivienda_GotFocus()
  MarcarTexto TxtVivienda
End Sub

Private Sub TxtVivienda_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtVivienda_LostFocus()
  TextoValido TxtVivienda, True, , 2
End Sub

Public Sub Importar_Rol_Pagos(Tipo_Rubro As String, OpcCosta As Boolean)
Dim FechaDec4Ini As String
Dim FechaDec4Fin As String
Dim FechaDec3Ini As String
Dim FechaDec3Fin As String
Dim Dias As Integer
  
  Anio = Format(Year(FechaSistema), "0000")
  FechaIni = BuscarFecha("01/01/" & Anio)
  FechaFin = BuscarFecha("31/12/" & Anio)
  
  Datos_IESS FechaFin
  
 'Fechas Decimo Cuarto
  If OpcCosta Then
     FechaDec4Ini = BuscarFecha("01/03/" & CStr(Val(Anio)) - 1)
     Mifecha = UltimoDiaMes("01/02/" & Anio)
     FechaDec4Fin = BuscarFecha(Mifecha)
  Else
     FechaDec4Ini = BuscarFecha("01/08/" & CStr(Val(Anio)) - 1)
     Mifecha = UltimoDiaMes("01/07/" & Anio)
     FechaDec4Fin = BuscarFecha(Mifecha)
  End If
  
 'fFechas Decimo Tercero
  FechaDec3Ini = BuscarFecha("01/12/" & CStr(Val(Anio)) - 1)
  Mifecha = UltimoDiaMes("01/11/" & Anio)
  FechaDec3Fin = BuscarFecha(Mifecha)

  If Tipo_Rubro = "Decimo_IV" Then
     sSQL = "UPDATE Catalogo_Rol_Pagos " _
          & "SET Valor_Dec_4to = 0, Dias_Dec_4to = 0 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Ejecutar_SQL_SP sSQL
  
     sSQL = "SELECT Codigo, COUNT(Codigo) As No_Meses, MIN(Fecha_D) As Fecha_Min " _
          & "FROM Trans_Rol_de_Pagos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Cod_Rol_Pago = '" & Tipo_Rubro & "' " _
          & "AND Fecha_D BETWEEN #" & FechaDec4Ini & "# AND #" & FechaDec4Fin & "# " _
          & "AND Ingresos > 0 " _
          & "GROUP BY Codigo "
     Select_Adodc AdoAux, sSQL
  Else
     sSQL = "UPDATE Catalogo_Rol_Pagos " _
          & "SET Valor_Dec_3ro = 0, Dias_Dec_3ro = 0 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "SELECT Codigo, COUNT(Codigo) As No_Meses, MIN(Fecha_D) As Fecha_Min " _
          & "FROM Trans_Rol_de_Pagos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Cod_Rol_Pago = '" & Tipo_Rubro & "' " _
          & "AND Fecha_D BETWEEN #" & FechaDec3Ini & "# AND #" & FechaDec3Fin & "# " _
          & "AND Ingresos > 0 " _
          & "GROUP BY Codigo "
     Select_Adodc AdoAux, sSQL
  End If
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Codigo = .fields("Codigo")
          Mifecha = .fields("Fecha_Min")
          FechaTexto = Mifecha
          sSQL = "SELECT Fecha " _
               & "FROM Catalogo_Rol_Pagos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo = '" & Codigo & "' " _
               & "AND Codigo = '" & Codigo & "' "
          Select_Adodc AdoDetalle, sSQL
          If AdoDetalle.Recordset.RecordCount > 0 Then FechaTexto = AdoDetalle.Recordset.fields("Fecha")
          Dias = CFechaLong(Mifecha) - CFechaLong(FechaTexto)
          If Dias < 365 Then Dias = (.fields("No_Meses") * 30) - Day(FechaTexto) Else Dias = .fields("No_Meses") * 30
          
          Total = Redondear((Sueldo_Basico / 360) * Dias, 2)
          If Tipo_Rubro = "Decimo_IV" Then
             sSQL = "UPDATE Catalogo_Rol_Pagos " _
                  & "SET Valor_Dec_4to = " & Total & ", Dias_Dec_4to = " & Dias & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Codigo = '" & Codigo & "' "
             Ejecutar_SQL_SP sSQL
          Else
             sSQL = "UPDATE Catalogo_Rol_Pagos " _
                  & "SET Valor_Dec_3ro = " & Total & ", Dias_Dec_3ro = " & Dias & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Codigo = '" & Codigo & "' "
             Ejecutar_SQL_SP sSQL
          End If
         .MoveNext
       Loop
   End If
  End With
  MsgBox "Importacion del " & Tipo_Rubro & " Exitoso, " & vbCrLf & "vuelva a ingresar para verificar los datos"
  Unload FRolPago
End Sub

Public Sub Leer_Catalogo_Rol_Pagos(GrupoRol As String)
    MBCta_Diferencia = FormatoCodigoCta("0")
    MBCta_Vacacion = FormatoCodigoCta("0")
    MBCta_Sueldo = FormatoCodigoCta("0")
    MBCta_Horas_Ext = FormatoCodigoCta("0")
    MBCta_Antig = FormatoCodigoCta("0")
    MBCta_Aporte_Patronal_G = FormatoCodigoCta("0")
    MBCta_Decimo_Cuarto_G = FormatoCodigoCta("0")
    MBCta_Decimo_Cuarto_P = FormatoCodigoCta("0")
    MBCta_Decimo_Tercer_G = FormatoCodigoCta("0")
    MBCta_Decimo_Tercer_P = FormatoCodigoCta("0")
    MBCta_Fondo_Reserva_G = FormatoCodigoCta("0")
    MBCta_Fondo_Reserva_P = FormatoCodigoCta("0")
    MBCta_Vacaciones_G = FormatoCodigoCta("0")
    MBCta_Vacaciones_P = FormatoCodigoCta("0")
    MBCta_IESS_Patronal = FormatoCodigoCta("0")
    MBCta_IESS_Personal = FormatoCodigoCta("0")
    MBCta_Quincena = FormatoCodigoCta("0")
    MBCta_ExtConyugue = FormatoCodigoCta("0")
    
    If Len(GrupoRol) > 1 Then
       With AdoGrupos.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Grupo_Rol = '" & GrupoRol & "' ")
            If Not .EOF Then
               MBCta_Diferencia = FormatoCodigoCta(.fields("Cta_Diferencia"))
               MBCta_Vacacion = FormatoCodigoCta(.fields("Cta_Vacacion"))
               MBCta_Sueldo = FormatoCodigoCta(.fields("Cta_Sueldo"))
               MBCta_Horas_Ext = FormatoCodigoCta(.fields("Cta_Horas_Ext"))
               MBCta_Antig = FormatoCodigoCta(.fields("Cta_Antiguedad"))
               MBCta_Aporte_Patronal_G = FormatoCodigoCta(.fields("Cta_Aporte_Patronal_G"))
               MBCta_Decimo_Cuarto_G = FormatoCodigoCta(.fields("Cta_Decimo_Cuarto_G"))
               MBCta_Decimo_Cuarto_P = FormatoCodigoCta(.fields("Cta_Decimo_Cuarto_P"))
               MBCta_Decimo_Tercer_G = FormatoCodigoCta(.fields("Cta_Decimo_Tercer_G"))
               MBCta_Decimo_Tercer_P = FormatoCodigoCta(.fields("Cta_Decimo_Tercer_P"))
               MBCta_Fondo_Reserva_G = FormatoCodigoCta(.fields("Cta_Fondo_Reserva_G"))
               MBCta_Fondo_Reserva_P = FormatoCodigoCta(.fields("Cta_Fondo_Reserva_P"))
               MBCta_Vacaciones_G = FormatoCodigoCta(.fields("Cta_Vacaciones_G"))
               MBCta_Vacaciones_P = FormatoCodigoCta(.fields("Cta_Vacaciones_P"))
               MBCta_IESS_Patronal = FormatoCodigoCta(.fields("Cta_IESS_Patronal"))
               MBCta_IESS_Personal = FormatoCodigoCta(.fields("Cta_IESS_Personal"))
               MBCta_Quincena = FormatoCodigoCta(.fields("Cta_Quincena"))
               MBCta_ExtConyugue = FormatoCodigoCta(.fields("Cta_Ext_Conyugue_P"))
            End If
        End If
       End With
    End If
End Sub

