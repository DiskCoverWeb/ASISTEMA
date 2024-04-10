VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form CrearEmp 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   Icon            =   "CrearEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   135
      Top             =   0
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Empresa"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Actualiza los Datos de la Empresa"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Produccion"
            Object.ToolTipText     =   "Encera los documentos electronicos de prueba"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin InetCtlsObjects.Inet URLInet 
      Left            =   8295
      Top             =   8190
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Left            =   105
      TabIndex        =   0
      Top             =   1155
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   12144
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "CrearEmp.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label24"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label20"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label35"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label34"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label16"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label27"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label32"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label33"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label21"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label44"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label38"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label31"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label15"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label30"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TextTelefono1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TextRUC"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TextEmpresa"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TextDireccion"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TextS_M"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TextGerente"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TextSubDir"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtCI"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtComercial"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TextFAX"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtContador"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtRUCCont"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtEmail"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "CProvincia"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "CNacion"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "CCiudadS"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "TextAbreviatura"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TxtEstablecimientos"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TxtRazonSocial"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtEmailContador"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TextTelefono2"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TxtNumPatronal"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "TxtCodBanco"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "TxtTipoCarga"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TxtEmailRespaldo"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtSeguro"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "CObligado"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "TxtSeguro2"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "TxtRUCOperadora"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "PROCESOS GENERALES"
      TabPicture(1)   =   "CrearEmp.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "LabelUsuario"
      Tab(1).Control(2)=   "LabelClave"
      Tab(1).Control(3)=   "DCListEmpCopy"
      Tab(1).Control(4)=   "TextLogoTipo"
      Tab(1).Control(5)=   "Picture1"
      Tab(1).Control(6)=   "File2"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(9)=   "Frame3"
      Tab(1).Control(10)=   "Frame5"
      Tab(1).Control(11)=   "CheqUsuario"
      Tab(1).Control(12)=   "CheqCopiiarEmpresa"
      Tab(1).Control(13)=   "TextUsuario"
      Tab(1).Control(14)=   "TextClave"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "COMPROBANTES ELECTRONICOS"
      TabPicture(2)   =   "CrearEmp.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      Begin VB.TextBox TxtRUCOperadora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10290
         MaxLength       =   13
         TabIndex        =   57
         Top             =   6405
         Width           =   1800
      End
      Begin VB.TextBox TextClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -65655
         MaxLength       =   10
         TabIndex        =   107
         Top             =   5565
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox TextUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -65655
         MaxLength       =   10
         TabIndex        =   105
         Top             =   5145
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.CheckBox CheqCopiiarEmpresa 
         Caption         =   "COPIAR SETEOS DE OTRA EMPRESA"
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
         Left            =   -70170
         TabIndex        =   108
         Top             =   5985
         Width           =   7260
      End
      Begin VB.CheckBox CheqUsuario 
         Caption         =   "ASIGNA USUARIO Y CLAVE DEL REPRESENTANTE LEGAL"
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
         Left            =   -70170
         TabIndex        =   103
         Top             =   5145
         Width           =   3270
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "| Servidor de Correos |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1695
         Left            =   -70170
         TabIndex        =   92
         Top             =   3255
         Width           =   7260
         Begin VB.CheckBox CheqConCopia 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Enviar Copia de Comprobantes por Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1890
            TabIndex        =   100
            Top             =   945
            Width           =   4005
         End
         Begin VB.TextBox TxtEmailProcesos 
            BeginProperty Font 
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
            TabIndex        =   101
            Text            =   "@"
            Top             =   1260
            Width           =   5790
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000C0C0&
            Height          =   645
            Left            =   5985
            Picture         =   "CrearEmp.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   945
            Width           =   1170
         End
         Begin VB.CheckBox CheqSecure 
            BackColor       =   &H00C0FFFF&
            Caption         =   " SECURE"
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
            Left            =   5985
            TabIndex        =   98
            Top             =   630
            Width           =   1170
         End
         Begin VB.CheckBox CheqSSL 
            BackColor       =   &H00C0FFFF&
            Caption         =   "SSL"
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
            Left            =   5985
            TabIndex        =   97
            Top             =   315
            Width           =   750
         End
         Begin VB.CheckBox CheqAutentificacion 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Autentificacion"
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
            TabIndex        =   99
            Top             =   945
            Width           =   1695
         End
         Begin VB.TextBox TxtPuerto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4935
            MaxLength       =   20
            TabIndex        =   96
            Top             =   525
            Width           =   960
         End
         Begin VB.TextBox TxtServidorSMTP 
            BeginProperty Font 
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
            TabIndex        =   94
            Top             =   525
            Width           =   4740
         End
         Begin VB.Label Label48 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PUERTO"
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
            Left            =   4935
            TabIndex        =   95
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SERVIDOR SMTP"
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
            Top             =   315
            Width           =   4740
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "| Cantidad de Decimales en |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   960
         Left            =   -74895
         TabIndex        =   79
         Top             =   5670
         Width           =   4635
         Begin VB.TextBox TxtDecPVP 
            BeginProperty Font 
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
            MaxLength       =   5
            TabIndex        =   81
            Top             =   525
            Width           =   1065
         End
         Begin VB.TextBox TxtDecCosto 
            BeginProperty Font 
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
            MaxLength       =   5
            TabIndex        =   83
            Top             =   525
            Width           =   960
         End
         Begin VB.TextBox TxtDecIVA 
            BeginProperty Font 
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
            MaxLength       =   5
            TabIndex        =   85
            Top             =   525
            Width           =   960
         End
         Begin VB.TextBox TxtDecCant 
            BeginProperty Font 
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
            MaxLength       =   5
            TabIndex        =   87
            Top             =   525
            Width           =   1170
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " P.V.P."
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
            TabIndex        =   80
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " COSTOS"
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
            TabIndex        =   82
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0FFC0&
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
            Height          =   225
            Left            =   2310
            TabIndex        =   84
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CANTIDAD"
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
            Left            =   3360
            TabIndex        =   86
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "| Numeración de Comprobantes |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1905
         Left            =   -74895
         TabIndex        =   68
         Top             =   3675
         Width           =   4635
         Begin VB.CheckBox OpcND 
            Caption         =   "N/D Secuenciales"
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
            Left            =   2205
            TabIndex        =   76
            Top             =   1260
            Width           =   2220
         End
         Begin VB.CheckBox OpcNDM 
            Caption         =   "N/D por meses"
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
            TabIndex        =   75
            Top             =   1260
            Width           =   2010
         End
         Begin VB.CheckBox OpcNC 
            Caption         =   "N/C Secuenciales"
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
            Left            =   2205
            TabIndex        =   78
            Top             =   1575
            Width           =   2220
         End
         Begin VB.CheckBox OpcNCM 
            Caption         =   "N/C por meses"
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
            TabIndex        =   77
            Top             =   1575
            Width           =   2010
         End
         Begin VB.CheckBox OpcCDM 
            Caption         =   "Diarios por meses"
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
            Width           =   2010
         End
         Begin VB.CheckBox OpcCIM 
            Caption         =   "Ingresos por meses"
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
            TabIndex        =   71
            Top             =   630
            Width           =   2010
         End
         Begin VB.CheckBox OpcCEM 
            Caption         =   "Egresos por meses"
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
            Top             =   945
            Width           =   2010
         End
         Begin VB.CheckBox OpcCD 
            Caption         =   "Diarios secuenciales"
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
            Left            =   2205
            TabIndex        =   70
            Top             =   315
            Width           =   2220
         End
         Begin VB.CheckBox OpcCI 
            Caption         =   "Ingresos secuenciales"
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
            Left            =   2205
            TabIndex        =   72
            Top             =   630
            Width           =   2220
         End
         Begin VB.CheckBox OpcCE 
            Caption         =   "Egresos secuenciales"
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
            Left            =   2205
            TabIndex        =   74
            Top             =   945
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "| Seteos Generales |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3270
         Left            =   -74895
         TabIndex        =   58
         Top             =   420
         Width           =   4635
         Begin VB.CheckBox CheqRegistrarIVA 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Registrar el IVA en el Asiento Contable"
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
            TabIndex        =   66
            Top             =   2520
            Width           =   3795
         End
         Begin VB.CheckBox CheqDetComp 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Procesar Detalle Auxiliar de Comprobantes"
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
            TabIndex        =   65
            Top             =   2205
            Width           =   4005
         End
         Begin VB.CheckBox CheqSuc 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Funciona como Matriz de Sucursales"
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
            TabIndex        =   67
            Top             =   2835
            Width           =   3585
         End
         Begin VB.CheckBox Cheq2Pag 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimir dos Roles Individuales por pagina"
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
            TabIndex        =   64
            Top             =   1890
            Width           =   4005
         End
         Begin VB.CheckBox CheqMedioRol 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimir Medio Rol"
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
            TabIndex        =   63
            Top             =   1575
            Width           =   2220
         End
         Begin VB.CheckBox CheqModFact 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Modificar Facturas o Notas de Venta"
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
            TabIndex        =   60
            Top             =   630
            Width           =   3585
         End
         Begin VB.CheckBox CheqModPVP 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Modificar Precio de Venta al Publico"
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
            TabIndex        =   61
            Top             =   945
            Width           =   3585
         End
         Begin VB.CheckBox CheqRecibo 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimir Recibo de Caja en Facturacion"
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
            TabIndex        =   62
            Top             =   1260
            Width           =   3900
         End
         Begin VB.CheckBox CheqSubMod 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Agrupar Saldos Detalle Auxiliar de Submodulos"
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
            TabIndex        =   59
            Top             =   315
            Width           =   4320
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "| Firma Electronica |"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   6420
         Left            =   -74895
         TabIndex        =   110
         Top             =   315
         Width           =   11985
         Begin VB.TextBox TxtPasword 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   126
            Top             =   2730
            Width           =   2640
         End
         Begin VB.TextBox TxtEmailConexion 
            BeginProperty Font 
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
            TabIndex        =   124
            Text            =   "@"
            Top             =   2730
            Width           =   9045
         End
         Begin VB.TextBox TxtPaswordCE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   130
            Top             =   3360
            Width           =   2640
         End
         Begin VB.TextBox TxtEmailConexionCE 
            BeginProperty Font 
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
            TabIndex        =   128
            Text            =   "@"
            Top             =   3360
            Width           =   9045
         End
         Begin VB.TextBox TxtContEsp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9135
            MaxLength       =   20
            TabIndex        =   114
            Top             =   210
            Width           =   2745
         End
         Begin VB.OptionButton OpcProduccion 
            Caption         =   "Ambiente en Produccion"
            BeginProperty Font 
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
            TabIndex        =   112
            Top             =   210
            Width           =   2535
         End
         Begin VB.OptionButton OpcAmbiente 
            Caption         =   "Ambiente de Prueba"
            BeginProperty Font 
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
            TabIndex        =   111
            Top             =   210
            Value           =   -1  'True
            Width           =   2220
         End
         Begin VB.TextBox TxtCertificado 
            BeginProperty Font 
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
            TabIndex        =   120
            Top             =   2100
            Width           =   9045
         End
         Begin VB.TextBox TxtPwdCertificado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   122
            Top             =   2100
            Width           =   2640
         End
         Begin VB.TextBox TxtWebAutorizacion 
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
            Height          =   330
            Left            =   105
            TabIndex        =   118
            Text            =   "."
            Top             =   1470
            Width           =   11775
         End
         Begin VB.TextBox TxtWebRecepcion 
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
            Height          =   330
            Left            =   105
            TabIndex        =   116
            Text            =   "."
            Top             =   840
            Width           =   11775
         End
         Begin VB.TextBox TxtLeyendaFA 
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
            Left            =   105
            MultiLine       =   -1  'True
            TabIndex        =   132
            Text            =   "CrearEmp.frx":132C
            Top             =   3990
            Width           =   11775
         End
         Begin VB.TextBox TxtLeyendaFA1 
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
            MultiLine       =   -1  'True
            TabIndex        =   134
            Text            =   "CrearEmp.frx":1461
            Top             =   5355
            Width           =   11775
         End
         Begin VB.Label Label45 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LEYENDA AL FINAL DE LA IMPRESION EN LA IMPRESORA DE PUNTO DE VENTA DE DOCUMENTOS ELECTRONICOS"
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
            TabIndex        =   133
            Top             =   5145
            Width           =   11775
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LEYENDA AL FINAL DE LOS DOCUMENTOS ELECTRONICOS"
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
            TabIndex        =   131
            Top             =   3780
            Width           =   11775
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CONTRASEÑA:"
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
            Left            =   9240
            TabIndex        =   129
            Top             =   3150
            Width           =   2640
         End
         Begin VB.Label Label37 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CONTRASEÑA:"
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
            Left            =   9240
            TabIndex        =   125
            Top             =   2520
            Width           =   2640
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMAIL PARA PROCESOS GENERALES:"
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
            TabIndex        =   123
            Top             =   2520
            Width           =   9045
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMAIL PARA DOCUMENTOS ELECTRONICOS:"
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
            TabIndex        =   127
            Top             =   3150
            Width           =   9045
         End
         Begin VB.Label Label43 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CONTRIBUYENTE ESPECIAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6405
            TabIndex        =   113
            Top             =   210
            Width           =   2745
         End
         Begin VB.Label Label42 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CONTRASEÑA:"
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
            Left            =   9240
            TabIndex        =   121
            Top             =   1890
            Width           =   2640
         End
         Begin VB.Label Label41 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CERTIFICADO FIRMA ELECTRONICA (DEBE SER EN FORMATO DE EXTENSION P12"
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
            TabIndex        =   119
            Top             =   1890
            Width           =   9045
         End
         Begin VB.Label Label40 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " WEBSERVICE SRI AUTORIZACION"
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
            TabIndex        =   117
            Top             =   1260
            Width           =   11775
         End
         Begin VB.Label Label39 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " WEBSERVICE SRI RECEPCION"
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
            TabIndex        =   115
            Top             =   630
            Width           =   11775
         End
      End
      Begin VB.TextBox TxtSeguro2 
         BeginProperty Font 
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
         MaxLength       =   5
         TabIndex        =   49
         Top             =   5775
         Width           =   1380
      End
      Begin VB.ComboBox CObligado 
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
         Left            =   1785
         TabIndex        =   10
         Top             =   1995
         Width           =   750
      End
      Begin VB.FileListBox File2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   -70170
         TabIndex        =   90
         Top             =   1050
         Width           =   2115
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   2640
         Left            =   -67965
         ScaleHeight     =   2580
         ScaleWidth      =   4995
         TabIndex        =   91
         Top             =   420
         Width           =   5055
      End
      Begin VB.TextBox TextLogoTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70170
         MaxLength       =   10
         TabIndex        =   89
         Text            =   "XXXXXXXXXX"
         Top             =   735
         Width           =   2115
      End
      Begin VB.TextBox TxtSeguro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   48
         Top             =   5775
         Width           =   1275
      End
      Begin VB.TextBox TxtEmailRespaldo 
         BeginProperty Font 
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
         TabIndex        =   46
         Text            =   "@"
         Top             =   5790
         Width           =   7365
      End
      Begin VB.TextBox TxtTipoCarga 
         BeginProperty Font 
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
         MaxLength       =   5
         TabIndex        =   38
         Top             =   3900
         Width           =   1065
      End
      Begin VB.TextBox TxtCodBanco 
         BeginProperty Font 
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
         MaxLength       =   5
         TabIndex        =   36
         Top             =   3900
         Width           =   1380
      End
      Begin VB.TextBox TxtNumPatronal 
         BeginProperty Font 
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
         MaxLength       =   15
         TabIndex        =   34
         Top             =   3900
         Width           =   1695
      End
      Begin VB.TextBox TextTelefono2 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   30
         Top             =   3900
         Width           =   1590
      End
      Begin VB.TextBox TxtEmailContador 
         BeginProperty Font 
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
         TabIndex        =   44
         Text            =   "@"
         Top             =   5160
         Width           =   11985
      End
      Begin VB.TextBox TxtRazonSocial 
         BeginProperty Font 
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
         MaxLength       =   120
         TabIndex        =   4
         Top             =   855
         Width           =   9885
      End
      Begin VB.TextBox TxtEstablecimientos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11340
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "000"
         Top             =   3270
         Width           =   750
      End
      Begin VB.TextBox TextAbreviatura 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10605
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "Ninguna"
         Top             =   3900
         Width           =   1485
      End
      Begin VB.ComboBox CCiudadS 
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
         Left            =   7455
         TabIndex        =   20
         Top             =   2640
         Width           =   4635
      End
      Begin VB.ComboBox CNacion 
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
         Left            =   105
         TabIndex        =   16
         Text            =   "ECUADOR"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox CProvincia 
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
         Left            =   2730
         TabIndex        =   18
         Text            =   "PICHINCHA"
         Top             =   2640
         Width           =   4635
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
         Left            =   105
         MaxLength       =   60
         TabIndex        =   42
         Text            =   "@"
         Top             =   4530
         Width           =   11985
      End
      Begin VB.TextBox TxtRUCCont 
         BeginProperty Font 
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
         MaxLength       =   13
         TabIndex        =   55
         Top             =   6420
         Width           =   1800
      End
      Begin VB.TextBox TxtContador 
         BeginProperty Font 
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
         TabIndex        =   53
         Top             =   6420
         Width           =   8205
      End
      Begin VB.TextBox TextFAX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3900
         Width           =   1590
      End
      Begin VB.TextBox TxtComercial 
         BeginProperty Font 
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
         MaxLength       =   120
         TabIndex        =   6
         Top             =   1275
         Width           =   9885
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
         Left            =   10395
         MaxLength       =   13
         TabIndex        =   14
         Top             =   1995
         Width           =   1695
      End
      Begin VB.TextBox TextSubDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10290
         MaxLength       =   10
         TabIndex        =   51
         Top             =   5790
         Width           =   1800
      End
      Begin VB.TextBox TextGerente 
         BeginProperty Font 
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
         MaxLength       =   60
         TabIndex        =   12
         Top             =   1995
         Width           =   7680
      End
      Begin VB.TextBox TextS_M 
         BeginProperty Font 
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
         MaxLength       =   5
         TabIndex        =   32
         Top             =   3900
         Width           =   960
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
         Left            =   105
         MaxLength       =   60
         TabIndex        =   22
         Top             =   3270
         Width           =   11145
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
         Left            =   2205
         MaxLength       =   80
         TabIndex        =   2
         Top             =   435
         Width           =   9885
      End
      Begin VB.TextBox TextRUC 
         BeginProperty Font 
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
         MaxLength       =   13
         TabIndex        =   8
         Top             =   2010
         Width           =   1590
      End
      Begin VB.TextBox TextTelefono1 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   26
         Top             =   3900
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo DCListEmpCopy 
         Bindings        =   "CrearEmp.frx":14E1
         DataSource      =   "AdoListEmpCopy"
         Height          =   360
         Left            =   -69855
         TabIndex        =   109
         Top             =   6300
         Visible         =   0   'False
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         Text            =   "Empresa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RUC OPERADORA"
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
         Left            =   10290
         TabIndex        =   56
         Top             =   6195
         Width           =   1800
      End
      Begin VB.Label LabelClave 
         Caption         =   "CLAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -66705
         TabIndex        =   106
         Top             =   5565
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label LabelUsuario 
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -66705
         TabIndex        =   104
         Top             =   5145
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OBLIG"
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
         Left            =   1785
         TabIndex        =   9
         Top             =   1785
         Width           =   750
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LOGO TIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70170
         TabIndex        =   88
         Top             =   420
         Width           =   2115
      End
      Begin VB.Label Label31 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SEGURO DESGRAVAMEN %"
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
         TabIndex        =   47
         Top             =   5580
         Width           =   2640
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL DE RESPALDO:"
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
         TabIndex        =   45
         Top             =   5580
         Width           =   7365
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO CAR."
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
         Left            =   9450
         TabIndex        =   37
         Top             =   3690
         Width           =   1065
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " COD. BANCO"
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
         Left            =   7980
         TabIndex        =   35
         Top             =   3690
         Width           =   1380
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NO. PATRONAL:"
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
         Left            =   6195
         TabIndex        =   33
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO 2:"
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
         Left            =   1785
         TabIndex        =   27
         Top             =   3690
         Width           =   1590
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL DE CONTABILIDAD:"
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
         TabIndex        =   43
         Top             =   4950
         Width           =   11985
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RAZON SOCIAL:"
         BeginProperty Font 
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
         Top             =   855
         Width           =   2115
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ESTA."
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
         Left            =   11340
         TabIndex        =   23
         Top             =   3060
         Width           =   750
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ABREVIATURA"
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
         Left            =   10605
         TabIndex        =   39
         Top             =   3690
         Width           =   1485
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NACIONALIDAD"
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
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   2430
         Width           =   2535
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVINCIA"
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
         Height          =   225
         Left            =   2730
         TabIndex        =   17
         Top             =   2430
         Width           =   4635
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CIUDAD"
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
         Height          =   225
         Left            =   7455
         TabIndex        =   19
         Top             =   2430
         Width           =   4635
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL DE LA EMPRESA:"
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
         Top             =   4320
         Width           =   11985
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RUC CONTADOR:"
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
         Left            =   8400
         TabIndex        =   54
         Top             =   6210
         Width           =   1800
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE DEL CONTADOR"
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
         TabIndex        =   52
         Top             =   6210
         Width           =   8205
      End
      Begin VB.Label Label6 
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
         Left            =   3465
         TabIndex        =   29
         Top             =   3690
         Width           =   1590
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE COMERCIAL:"
         BeginProperty Font 
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
         TabIndex        =   5
         Top             =   1275
         Width           =   2115
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " C.I./ PASAPORTE"
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
         Left            =   10395
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " REPRESENTANTE LEGAL:"
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
         Left            =   2625
         TabIndex        =   11
         Top             =   1800
         Width           =   7680
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DIRECCION MATRIZ:"
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
         TabIndex        =   21
         Top             =   3060
         Width           =   11145
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SUBDIR.:"
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
         Left            =   10290
         TabIndex        =   50
         Top             =   5580
         Width           =   1800
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MONEDA"
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
         Left            =   5145
         TabIndex        =   31
         Top             =   3690
         Width           =   960
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   435
         Width           =   2115
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RUC:"
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
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   105
         TabIndex        =   25
         Top             =   3690
         Width           =   1590
      End
   End
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   435
      Left            =   315
      Top             =   2730
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   767
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
      Caption         =   "Empresa"
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
   Begin MSDataListLib.DataCombo DCListEmp 
      Bindings        =   "CrearEmp.frx":14FE
      DataSource      =   "AdoListEmp"
      Height          =   360
      Left            =   2520
      TabIndex        =   137
      Top             =   735
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Empresa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoListEmp 
      Height          =   330
      Left            =   315
      Top             =   1995
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
      Caption         =   "ListEmp"
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
   Begin MSAdodcLib.Adodc AdoBusqEmp 
      Height          =   330
      Left            =   315
      Top             =   2415
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
      Caption         =   "BusqEmp"
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
   Begin MSAdodcLib.Adodc AdoPaises 
      Height          =   330
      Left            =   315
      Top             =   3255
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
      Caption         =   "Paises"
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
   Begin MSAdodcLib.Adodc AdoMySQLClave 
      Height          =   330
      Left            =   315
      Top             =   3675
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "MySQLClave"
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
   Begin MSAdodcLib.Adodc AdoClave 
      Height          =   330
      Left            =   315
      Top             =   4095
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Clave"
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
   Begin MSAdodcLib.Adodc AdoListEmpCopy 
      Height          =   330
      Left            =   315
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
      Caption         =   "ListEmpCopy"
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
      Left            =   7665
      Top             =   8190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CrearEmp.frx":1517
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CrearEmp.frx":1831
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CrearEmp.frx":1B4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CrearEmp.frx":1E65
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LISTAS DE EMPRESAS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   136
      Top             =   735
      Width           =   2430
   End
End
Attribute VB_Name = "CrearEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SubDirOld As String
Dim TipoBenefCI As String
Dim NTabla As String
Dim vLeyendaFA As String
Dim vLeyendaFA1 As String

'''Public Sub Listar_Tablas()
'''Dim AdoCon1 As ADODB.Connection
'''Dim RstSchema As ADODB.Recordset
'''Dim IJ As Long
'''Dim IdTime As Long
'''Dim strCnn As String
'''' Consultamos las cuentas de la tabla
'''  RatonReloj
'''  LstTablas.Clear
'''  Set AdoCon1 = New ADODB.Connection
'''  AdoCon1.open AdoStrCnn
'''  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
'''  Do Until RstSchema.EOF
'''     If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
'''        LstTablas.AddItem RstSchema!TABLE_NAME
'''     End If
'''     RstSchema.MoveNext
'''  Loop
'''  RatonNormal
'''End Sub

Private Sub CCiudadS_GotFocus()
  MarcarTexto CCiudadS
End Sub

Private Sub CCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CheqCopiiarEmpresa_Click()
    If CheqCopiiarEmpresa.value = 0 Then
       DCListEmpCopy.Visible = False
    Else
       sSQL = "SELECT Empresa, Item " _
            & "FROM Empresas " _
            & "WHERE Item <> '000' " _
            & "ORDER BY Empresa "
       SelectDB_Combo DCListEmpCopy, AdoListEmpCopy, sSQL, "Empresa"
       DCListEmpCopy.Visible = True
       DCListEmpCopy.SetFocus
    End If
End Sub

Private Sub CheqUsuario_Click()
    If CheqUsuario.value = 0 Then
        LabelUsuario.Visible = False
        LabelClave.Visible = False
        TextUsuario.Visible = False
        TextClave.Visible = False
    Else
        LabelUsuario.Visible = True
        LabelClave.Visible = True
        TextUsuario.Visible = True
        TextClave.Visible = True
        sSQL = "SELECT Codigo, Usuario, Clave, ID " _
             & "FROM Accesos " _
             & "WHERE Codigo = '" & TxtCI.Text & "' "
        Select_Adodc AdoClave, sSQL
        If AdoClave.Recordset.RecordCount > 0 Then
           TextUsuario.Text = AdoClave.Recordset.fields("Usuario")
           TextClave.Text = AdoClave.Recordset.fields("Clave")
           TextClave.SetFocus
        Else
           TextUsuario.Text = ""
           TextClave.Text = ""
           TextUsuario.SetFocus
        End If
    End If
End Sub

Private Sub CNacion_GotFocus()
  MarcarTexto CNacion
End Sub

Private Sub CNacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CNacion_LostFocus()
  CodigoPais = SinEspaciosIzq(CNacion)
  NombrePais = TrimStrg(MidStrg(CNacion, Len(CodigoPais) + 1, Len(CNacion)))
End Sub

'Grabar datos empresa
Private Sub Grabar_Empresa()
Dim PeriodoCopy As String
Dim Ind As Integer

  RatonReloj
  TextoValido TextEmpresa
  TextoValido TxtRazonSocial
  TextoValido TxtComercial
  TextoValido TextLogoTipo
  If Val(TxtDecPVP) < 2 Then TxtDecPVP = "2"
  If Val(TxtDecCosto) < 2 Then TxtDecCosto = "2"
  If Val(TxtDecIVA) < 2 Then TxtDecIVA = "2"
  If Val(TxtDecIVA) > 4 Then TxtDecIVA = "4"
  If Len(TxtEmailRespaldo) <= 1 Then TxtEmailRespaldo = CorreoDiskCover
  If Len(TxtEmailConexion) <= 1 Then TxtEmailConexion = CorreoDiskCover
  If Len(TxtPasword) <= 1 Then TxtPasword = ContrasenaDiskCover
  If Len(TxtLeyendaFA) <= 1 Then TxtLeyendaFA = Ninguno
  If Len(TxtLeyendaFA1) <= 1 Then TxtLeyendaFA1 = Ninguno
  
  CodigoPais = SinEspaciosIzq(CNacion)
  NombrePais = TrimStrg(MidStrg(CNacion, Len(CodigoPais) + 1, Len(CNacion)))
  TxtContador = ULCase(TxtContador)
  TextGerente = ULCase(TextGerente)
  TextDireccion = UCaseStrg(TextDireccion)
  TxtComercial = UCase(TxtComercial)
  
  Num_Meses_CD = False
  Num_Meses_CE = False
  Num_Meses_CI = False
  Num_Meses_ND = False
  Num_Meses_NC = False
  
  If OpcCDM.value <> 0 Then Num_Meses_CD = True
  If OpcCEM.value <> 0 Then Num_Meses_CE = True
  If OpcCIM.value <> 0 Then Num_Meses_CI = True
  If OpcNDM.value <> 0 Then Num_Meses_ND = True
  If OpcNCM.value <> 0 Then Num_Meses_NC = True

  With AdoListEmp.Recordset
   If .RecordCount > 0 Then
       If TxtCodBanco <> .fields("CodBanco") Then Control_Procesos Normal, "Modifico: Cod. Banco " & TxtCodBanco
       If TxtContador <> .fields("Contador") Then Control_Procesos Normal, "Modifico: Nom. Contador " & TxtContador
       If TextEmpresa <> .fields("Empresa") Then Control_Procesos Normal, "Modifico: Empresa " & TextEmpresa
       If TextGerente <> .fields("Gerente") Then Control_Procesos Normal, "Modifico: Gerente " & TextGerente
       If TextLogoTipo <> .fields("Logo_Tipo") Then Control_Procesos Normal, "Modifico: Logotipo " & TextLogoTipo
       If TextRUC <> .fields("RUC") Then Control_Procesos Normal, "Modifico: RUC " & TextRUC
       If TextTelefono1 <> .fields("Telefono1") Then Control_Procesos Normal, "Modifico: Telefono1 " & TextTelefono1
       If TextTelefono2 <> .fields("Telefono2") Then Control_Procesos Normal, "Modifico: Telefono2 " & TextTelefono2
       If TextFAX <> .fields("FAX") Then Control_Procesos Normal, "Modifico: FAX " & TextFAX
       If TextS_M <> .fields("S_M") Then Control_Procesos Normal, "Modifico: Modena " & TextS_M
       If CCiudadS <> .fields("Ciudad") Then Control_Procesos Normal, "Modifico: Ciudad " & CCiudadS
       If NombrePais <> .fields("Pais") Then Control_Procesos Normal, "Modifico: Pais " & NombrePais
       If TextDireccion <> .fields("Direccion") Then Control_Procesos Normal, "Modifico: Direccion " & TextDireccion
       If TextSubDir <> .fields("SubDir") Then Control_Procesos Normal, "Modifico: Sub Dir " & TextSubDir
       If TxtEmail <> .fields("Email") Then Control_Procesos Normal, "Modifico: Email " & TxtEmail
       If TxtRUCCont <> .fields("RUC_Contador") Then Control_Procesos Normal, "Modifico: RUC Cont. " & TxtRUCCont
       If TxtCI <> .fields("CI_Representante") Then Control_Procesos Normal, "Modifico: CI Represen " & TxtCI
       If TxtComercial <> .fields("Nombre_Comercial") Then Control_Procesos Normal, "Modifico: Nom. Comercial " & TxtComercial
   End If
  End With
  Numero = 0
  Si_No = True
  Encontro = True
  sSQL = "SELECT Item " _
       & "FROM Empresas " _
       & "WHERE Item <> '000' " _
       & "ORDER BY Item "
  Select_Adodc AdoBusqEmp, sSQL
  With AdoBusqEmp.Recordset
   If .RecordCount > 0 Then
       Numero = Val(.fields("Item"))
       For I = Numero To 999
          .MoveFirst
          .Find ("Item = '" & Format(I, "000") & "' ")
           If .EOF Then
               Numero = I
               Exit For
           End If
       Next I
   End If
  End With
  
  sSQL = "SELECT " & Full_Fields("Empresas") & " " _
       & "FROM Empresas " _
       & "WHERE Empresa = '" & TextEmpresa & "' " _
       & "AND RUC = '" & TextRUC & "' "
  Select_Adodc AdoEmpresa, sSQL
  With AdoEmpresa.Recordset
   If .RecordCount <= 0 Then
       Control_Procesos Normal, "Crear la Empresa " & UCaseStrg(TextEmpresa)
       OpcAmbiente.value = 1
       I = 1
       Cadena = MidStrg(TextEmpresa, 1, 2)
       Do While I <= Len(TextEmpresa)
          I = I + 1
          If MidStrg(TextEmpresa, I, 1) = " " Then
             If MidStrg(TextEmpresa, I + 1, 1) = " " Then I = I + 1 Else Cadena = Cadena & MidStrg(TextEmpresa, I + 1, 2)
          End If
       Loop
       NumEmpresa = Format$(Numero, "000")
       Cadena = Replace(Cadena, " ", "")
       Cadena = Replace(Cadena, ".", "")
       If Len(Cadena) > 10 Then Cadena = MidStrg(Cadena, 1, 10) Else Cadena = Cadena & String(10 - Len(Cadena), "_")
       TextAbreviatura = UCase(Cadena)
       If Len(Cadena) > 7 Then Cadena = MidStrg(Cadena, 1, 7) Else Cadena = Cadena & String(7 - Len(Cadena), "_")
       TextSubDir = UCase(Cadena) & NumEmpresa
       SetAddNew AdoEmpresa
       SetFields AdoEmpresa, "Item", NumEmpresa
       SetFields AdoEmpresa, "Grupo", NumEmpresa
       SetFields AdoEmpresa, "Fecha", FechaSistema
       SetFields AdoEmpresa, "Alto", 11
       SetFields AdoEmpresa, "Formato_Inventario", "CC.CC.CCC.CCCCCC"
       SetFields AdoEmpresa, "Formato_Activo", "CC.CC.CCC.CCCCCC"
       SetFields AdoEmpresa, "Formato_Cuentas", "C.C.CC.CC.CC.CCC"
'''       SetFields AdoEmpresa, "", ""
   Else
       NumEmpresa = .fields("Item")
       Si_No = False
       Control_Procesos Normal, "Modificar Datos de " & UCaseStrg(TextEmpresa)
   End If
   SetFields AdoEmpresa, "Obligado_Conta", CObligado.Text
   SetFields AdoEmpresa, "CProv", MidStrg(CProvincia.Text, 1, 2)
   SetFields AdoEmpresa, "CodBanco", TxtCodBanco
   SetFields AdoEmpresa, "Contador", TxtContador
   SetFields AdoEmpresa, "Logo_Tipo", TextLogoTipo
   SetFields AdoEmpresa, "Empresa", TextEmpresa
   SetFields AdoEmpresa, "Gerente", TextGerente
   SetFields AdoEmpresa, "RUC", TextRUC
   SetFields AdoEmpresa, "Telefono1", TextTelefono1
   SetFields AdoEmpresa, "Telefono2", TextTelefono2
   SetFields AdoEmpresa, "FAX", TextFAX
   SetFields AdoEmpresa, "S_M", TextS_M
   SetFields AdoEmpresa, "Ciudad", CCiudadS
   SetFields AdoEmpresa, "CPais", CodigoPais
   SetFields AdoEmpresa, "Pais", NombrePais
   SetFields AdoEmpresa, "Direccion", TextDireccion
   SetFields AdoEmpresa, "Email", TxtEmail
   
   SetFields AdoEmpresa, "Email_Contabilidad", TxtEmailContador
   SetFields AdoEmpresa, "Establecimientos", TxtEstablecimientos
   SetFields AdoEmpresa, "Num_CD", CBool(OpcCDM.value)
   SetFields AdoEmpresa, "Num_CE", CBool(OpcCEM.value)
   SetFields AdoEmpresa, "Num_CI", CBool(OpcCIM.value)
   SetFields AdoEmpresa, "Num_NC", CBool(OpcNCM.value)
   SetFields AdoEmpresa, "Num_ND", CBool(OpcNDM.value)
   SetFields AdoEmpresa, "Det_Comp", CBool(CheqDetComp.value)
   SetFields AdoEmpresa, "Mod_Fact", CBool(CheqModFact.value)
   SetFields AdoEmpresa, "Mod_PVP", CBool(CheqModPVP.value)
   SetFields AdoEmpresa, "Det_SubMod", CBool(CheqSubMod.value)
   SetFields AdoEmpresa, "Rol_2_Pagina", CBool(Cheq2Pag.value)
   SetFields AdoEmpresa, "Medio_Rol", CBool(CheqMedioRol.value)
   SetFields AdoEmpresa, "Registrar_IVA", CBool(CheqRegistrarIVA.value)
   SetFields AdoEmpresa, "Imp_Recibo_Caja", CBool(CheqRecibo.value)
   'SetFields AdoEmpresa, "No_ATS", CBool(CheqNoATS.value)
   SetFields AdoEmpresa, "RUC_Contador", TxtRUCCont
   SetFields AdoEmpresa, "CI_Representante", TxtCI
   SetFields AdoEmpresa, "Razon_Social", TxtRazonSocial
   SetFields AdoEmpresa, "Nombre_Comercial", TxtComercial
   SetFields AdoEmpresa, "Dec_PVP", Val(TxtDecPVP)
   SetFields AdoEmpresa, "Dec_Costo", Val(TxtDecCosto)
   SetFields AdoEmpresa, "Dec_IVA", Val(TxtDecIVA)
   SetFields AdoEmpresa, "Dec_Cant", Val(TxtDecCant)
   SetFields AdoEmpresa, "Seguro", Val(TxtSeguro)
   SetFields AdoEmpresa, "Seguro2", Val(TxtSeguro2)
   SetFields AdoEmpresa, "Email_Respaldos", TxtEmailRespaldo
   SetFields AdoEmpresa, "Email_Procesos", TxtEmailProcesos
   SetFields AdoEmpresa, "Email_Conexion", TxtEmailConexion
   SetFields AdoEmpresa, "Email_Contraseña", TxtPasword
   SetFields AdoEmpresa, "Email_Conexion_CE", TxtEmailConexionCE
   SetFields AdoEmpresa, "Email_Contraseña_CE", TxtPaswordCE
   SetFields AdoEmpresa, "LeyendaFA", TxtLeyendaFA
   SetFields AdoEmpresa, "LeyendaFAT", TxtLeyendaFA1
   SetFields AdoEmpresa, "Web_SRI_Recepcion", TxtWebRecepcion
   SetFields AdoEmpresa, "Web_SRI_Autorizado", TxtWebAutorizacion
   SetFields AdoEmpresa, "Ruta_Certificado", TxtCertificado
   SetFields AdoEmpresa, "Clave_Certificado", TxtPwdCertificado
   SetFields AdoEmpresa, "Codigo_Contribuyente_Especial", TxtContEsp
   SetFields AdoEmpresa, "Tipo_Carga_Banco", TxtTipoCarga
   If CheqConCopia.value = 1 Then SetFields AdoEmpresa, "Email_CE_Copia", True Else SetFields AdoEmpresa, "Email_CE_Copia", False
   If OpcAmbiente.value Then SetFields AdoEmpresa, "Ambiente", "1" Else SetFields AdoEmpresa, "Ambiente", "2"
   SetFields AdoEmpresa, "smtp_Servidor", TxtServidorSMTP
   SetFields AdoEmpresa, "smtp_Puerto", TxtPuerto.Text
   SetFields AdoEmpresa, "smtp_UseAuntentificacion", CBool(CheqAutentificacion.value)
   SetFields AdoEmpresa, "smtp_SSL", CBool(CheqSSL.value)
   SetFields AdoEmpresa, "smtp_Secure", CBool(CheqSecure.value)
   
   DigVerif = Digito_Verificador(TxtCI)
   SetFields AdoEmpresa, "TD", Tipo_RUC_CI.Tipo_Beneficiario
   SetFields AdoEmpresa, "SubDir", UCaseStrg(TextSubDir.Text)
   SetFields AdoEmpresa, "Abreviatura", TextAbreviatura.Text
   SetFields AdoEmpresa, "RUC_Operadora", TxtRUCOperadora.Text
   SetUpdate AdoEmpresa
  End With
   
   If CheqUsuario.value <> 0 Then
      sSQL = "SELECT Codigo, Usuario, Clave, ID " _
           & "FROM Accesos " _
           & "WHERE Codigo = '" & TxtCI.Text & "' "
      Select_Adodc AdoClave, sSQL
      If AdoClave.Recordset.RecordCount > 0 Then
         If Len(TextClave.Text) > 1 Then
            AdoClave.Recordset.fields("Clave") = TextClave.Text
            AdoClave.Recordset.Update
         End If
      Else
         Control_Procesos Normal, "Creacion de usuario: " & TextUsuario.Text:
         SetAdoAddNew "Accesos", True
         SetAdoFields "TODOS", True
         SetAdoFields "Clave", TextClave.Text
         SetAdoFields "Codigo", TxtCI.Text
         SetAdoFields "Usuario", TextUsuario.Text
         SetAdoFields "Nombre_Completo", TextGerente.Text
         SetAdoUpdate
       
        'Grabamos en Clientes Tambien
         sSQL = "SELECT Codigo " _
              & "FROM Clientes " _
              & "WHERE Codigo = '" & TxtCI.Text & "' "
         Select_Adodc AdoClave, sSQL
         If AdoClave.Recordset.RecordCount <= 0 Then
            SetAdoAddNew "Clientes", True
            SetAdoFields "Codigo", TxtCI.Text
            SetAdoFields "CI_RUC", TxtCI.Text
            SetAdoFields "TD", "C"
            SetAdoFields "Cliente", UCaseStrg(TextGerente.Text)
            SetAdoFields "Grupo", NumEmpresa
            SetAdoUpdate
         End If
      End If
      
     'Grabamos en MySQL el mismo usuario
      sSQL = "SELECT CI_NIC, Nombre_Usuario, Usuario, Clave, TODOS " _
           & "FROM acceso_usuarios " _
           & "WHERE CI_NIC = '" & TxtCI.Text & "' "
      Select_Adodc AdoMySQLClave, sSQL
      If AdoMySQLClave.Recordset.RecordCount <= 0 Then
         AdoMySQLClave.Recordset.AddNew
         AdoMySQLClave.Recordset.fields("CI_NIC") = TxtCI.Text
         AdoMySQLClave.Recordset.fields("Nombre_Usuario") = TextGerente.Text
         AdoMySQLClave.Recordset.fields("Usuario") = TextUsuario.Text
         AdoMySQLClave.Recordset.fields("Clave") = TextClave.Text
         AdoMySQLClave.Recordset.fields("TODOS") = True
         AdoMySQLClave.Recordset.Update
      Else
         If Len(TextClave.Text) > 1 Then
            AdoMySQLClave.Recordset.fields("Clave") = TextClave.Text
            AdoMySQLClave.Recordset.Update
         End If
      End If
   End If
  
  'Si es nueva empresa crea los codigo por defaul
   If Si_No Then
      If CheqCopiiarEmpresa.value <> 0 Then
         PeriodoCopy = Periodo_Contable
         Cadena = DCListEmpCopy.Text
         NumItem = Ninguno
         If Cadena = "" Then Cadena = Ninguno
         With AdoListEmpCopy.Recordset
          If .RecordCount > 0 Then
             .MoveFirst
             .Find ("Empresa LIKE '" & Cadena & "' ")
              If Not .EOF Then NumItem = .fields("Item")
          End If
         End With
         If NumItem <> Ninguno Then
            Copiar_Tabla_SP "Catalogo_Cuentas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
            Copiar_Tabla_SP "Codigos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
            Copiar_Tabla_SP "Ctas_Proceso", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
         
            Copiar_Tabla_SP "Catalogo_Lineas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
            Copiar_Tabla_SP "Catalogo_Productos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
         
            Copiar_Tabla_SP "Formato", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
            Copiar_Tabla_SP "Seteos_Documentos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
           
            Copiar_Tabla_SP "Catalogo_SubCtas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
            Copiar_Tabla_SP "Catalogo_CxCxP", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
          
            sSQL = "UPDATE Catalogo_Lineas " _
                 & "SET Autorizacion = '" & TextRUC & "', Codigo = 'S' + Serie " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Fact = 'FA' " _
                 & "AND LEN(Autorizacion) = 13 "
            Ejecutar_SQL_SP sSQL
         End If
      Else
         sSQL = "DELETE * " _
              & "FROM Codigos " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Ninguno & "' "
         Ejecutar_SQL_SP sSQL
        
         sSQL = "INSERT INTO Codigos (Item, Concepto, Numero, Periodo, X) " _
              & "SELECT '" & NumEmpresa & "', Concepto, Numero, Periodo, X " _
              & "FROM Codigos " _
              & "WHERE Item = '000' " _
              & "AND Periodo = '.' " _
              & "ORDER BY Concepto, Numero "
         Ejecutar_SQL_SP sSQL
      End If
   End If
  
  'Verificamos si esta creada la empresa
   RUC = TextRUC
   Empresa = TextEmpresa
   RazonSocial = TxtRazonSocial
   NombreCiudad = CCiudadS
   NombreContador = TxtContador
   RUC_Contador = TxtRUCCont
   NombreGerente = TextGerente
   NLogoTipo = TextLogoTipo
   NMarcaAgua = "DISKCOVER"
   CadenaParcial = ""

   If CheqUsuario.value <> 0 Then
      CodigoUsuario = TxtCI
      NombreUsuario = TextGerente
      IDEUsuario = TextUsuario
      PWRUsuario = TextClave
      EmailUsuario = TxtEmail
   Else
      CodigoUsuario = Ninguno
      NombreUsuario = Ninguno
      IDEUsuario = Ninguno
      PWRUsuario = Ninguno
      EmailUsuario = Ninguno
   End If
  '|--=:******* CONECCON A MYSQL *******:=--|
   Datos_Iniciales_Entidad_SP_MySQL
  '|--=:******* --------.------- *******:=--|
   RatonNormal
   Unload CrearEmp
   MsgBox "Proceso Terminado, vuelva a ingresar para verificar los cambios"
   End
  'ListEmp.Show
End Sub

Private Sub Eliminar()
  If ClaveSupervisor Then
     Mensajes = "Esta seguro de eliminar: " & TextEmpresa.Text
     Titulo = "Pregunta de Eliminación"
     If BoxMensaje = vbYes Then
        Eliminar_Empresa_SP NumEmpresa, TextEmpresa
        Evaluar = False
        Cadena = Dir(RutaSistema & "\EMPRESA\", vbDirectory) 'Recupera la primera entrada.
        Do While Cadena <> ""
           If Cadena <> "." And Cadena <> ".." Then
              If (GetAttr(RutaSistema & "\EMPRESA\" & Cadena) And vbDirectory) = vbDirectory Then
                 If UCaseStrg(Cadena) = UCaseStrg(TextSubDir) Then Evaluar = True
              End If
           End If
           Cadena = Dir
        Loop
        If Evaluar Then
           Si_No = False
           Cadena = Dir(RutaSistema & "\EMPRESA\" & TextSubDir & "\", vbNormal) 'Recupera la primera entrada.
           Do While Cadena <> ""
              If Cadena <> "." And Cadena <> ".." Then
                 If (GetAttr(RutaSistema & "\EMPRESA\" & TextSubDir & "\" & Cadena) And vbNormal) = vbNormal Then
                     Si_No = True
                 End If
              End If
              Cadena = Dir
           Loop
           If Si_No Then Kill RutaSistema & "\EMPRESA\" & TextSubDir & "\*.*"
           RmDir RutaSistema & "\EMPRESA\" & TextSubDir
        End If
        Unload CrearEmp
        ListEmp.Show 1
     End If
  End If
End Sub

Private Sub Command1_Click()
    RatonReloj
    If Len(TxtServidorSMTP.Text) > 1 And Val(TxtPuerto.Text) > 0 Then
       sSQL = "UPDATE Empresas " _
            & "SET smtp_Servidor = '" & TxtServidorSMTP.Text & "', smtp_Puerto = " & Val(TxtPuerto.Text) & " " _
            & "WHERE Item = '" & NumEmpresa & "' "
       Ejecutar_SQL_SP sSQL
    End If
    TMail.servidor = TxtServidorSMTP.Text                        ' smtp.gmail.com
    TMail.de = TxtEmailConexionCE.Text
    TMail.Usuario = TxtEmailConexionCE.Text
    TMail.Password = TxtPaswordCE.Text
    TMail.Puerto = TxtPuerto.Text                                ' 465
    TMail.useAuntentificacion = CBool(CheqAutentificacion.value) ' True
    TMail.ssl = CBool(CheqSSL.value)                             ' True
   '---------------------------------------------------------------
    NombreGerente = "Walter Vaca Prieto"
    RazonSocial = "DISKCOVER SYSTEM"
    TMail.Adjunto = ""
    TMail.para = ""
    Insertar_Mail TMail.para, CorreoDiskCover
    Insertar_Mail TMail.para, TxtEmailProcesos.Text
    Insertar_Mail TMail.para, TxtEmailConexionCE.Text
    'MsgBox TMail.para
    TMail.Asunto = "Prueba de envio correos desde DiskCover System"
    TMail.Mensaje = "Envio correo de prueba, para verificar que los parametros estan correctos, de la Empresa: " _
                  & TxtRazonSocial & vbCrLf _
                  & "SERVIRLES ES NUESTRO COMPROMISO, DISFRUTARLO ES EL SUYO."
    FEnviarCorreos.Show 1
    RatonNormal
    MsgBox "Verifique que llego el correo a su bandeja de entrada"
End Sub

Private Sub Encera_Comprobantes_Electronicos()
  If ClaveGerente Then
     RatonReloj
     sSQL = "DELETE * " _
          & "FROM Facturas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Autorizacion) >= 13 "
     Ejecutar_SQL_SP sSQL

     sSQL = "DELETE * " _
          & "FROM Detalle_Factura " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Autorizacion) >= 13 "
     Ejecutar_SQL_SP sSQL

     sSQL = "DELETE * " _
          & "FROM Trans_Abonos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Autorizacion) >= 13 "
     Ejecutar_SQL_SP sSQL

     sSQL = "UPDATE Codigos " _
          & "SET Numero = 1 " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Concepto LIKE '%_SERIE_%' "
     Ejecutar_SQL_SP sSQL
     RatonNormal
     MsgBox "Proceso Terminado, Procesa a Operar"
  End If
End Sub

Private Sub CProvincia_GotFocus()
  MarcarTexto CProvincia
End Sub

Private Sub CProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CProvincia_LostFocus()
  CCiudadS.Clear
  sSQL = "SELECT " & Full_Fields("Tabla_Naciones") & " " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "AND CProvincia = '" & SinEspaciosIzq(CProvincia) & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoPaises, sSQL
  If AdoPaises.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoPaises.Recordset.fields("Descripcion_Rubro")
     Do While Not AdoPaises.Recordset.EOF
        CCiudadS.AddItem AdoPaises.Recordset.fields("Descripcion_Rubro")
        AdoPaises.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
  CCiudadS.SetFocus
End Sub

Private Sub DCListEmp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCListEmp_LostFocus()
  LlenarEmpresa DCListEmp
End Sub

Private Sub DCListEmpCopy_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCListEmpCopy_LostFocus()
  SSTab1.Tab = 2
  OpcAmbiente.SetFocus
End Sub

Private Sub File2_Click()
Dim LogoTipo1 As String
  RatonReloj
  LogoTipo1 = RutaSistema & "\LOGOS\" & File2.Filename
  Picture1.Picture = LoadPicture()
  Picture1.AutoRedraw = True
  Picture1.PaintPicture LoadPicture(LogoTipo1), 0.01, 0.01, 5100, 2550
  TextLogoTipo.Text = UCaseStrg(ObtenerArchivo(File2.Filename))
  RatonNormal
End Sub

Private Sub File2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     Mensajes = "De verdad desea eliminar el archivo: " & File2.Filename
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = 6 Then
        Kill RutaSistema & "\LOGOS\" & File2.Filename
        File2.Refresh
     End If
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT " & Full_Fields("Empresas") & " " _
       & "FROM Empresas " _
       & "WHERE Item <> '000' " _
       & "ORDER BY Empresa "
  SelectDB_Combo DCListEmp, AdoListEmp, sSQL, "Empresa"

  vLeyendaFA = "Para consultas, requerimientos o reclamos puede contactarse a nuestro Centro de Atención al Cliente Teléfono: NUMERO_TELEFONO, " _
             & "o escriba al correo EMAIL_EMPRESA; para Transferencia o Depósitos hacer en El Banco NOMBRE_BANCO a " _
             & "Nombre de REPRESENTANTE_LEGAL/CTA_AHR_CTE_NUMERO, a Nombre de RAZON_SOCIAL"
             
  vLeyendaFA1 = "ESLOGAN_DE_LA_EMPRESA"
       
  CrearEmp.Caption = "CREACION/ELIMINACION/MODIFICACION DE EMPRESAS"
  Periodo_Contable = Ninguno
  
  File2.Filename = RutaSistema & "\LOGOS\*.jpg"
  
  CObligado.Clear
  CObligado.AddItem "NN"
  CObligado.AddItem "SI"
  CObligado.AddItem "NO"
  CObligado.Text = "NN"
  
  CNacion.Clear
  sSQL = "SELECT " & Full_Fields("Tabla_Naciones") & " " _
       & "FROM Tabla_Naciones " _
       & "WHERE TR = 'N' " _
       & "ORDER BY CPais,Descripcion_Rubro "
  Select_Adodc AdoPaises, sSQL
  If AdoPaises.Recordset.RecordCount > 0 Then
     CNacion.Text = "593 ECUADOR"
     Do While Not AdoPaises.Recordset.EOF
        CNacion.AddItem AdoPaises.Recordset.fields("CPais") & " " & AdoPaises.Recordset.fields("Descripcion_Rubro")
        AdoPaises.Recordset.MoveNext
     Loop
  End If
  CNacion.AddItem "999 OTRO"
  CProvincia.Clear
  sSQL = "SELECT " & Full_Fields("Tabla_Naciones") & " " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoPaises, sSQL
  If AdoPaises.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoPaises.Recordset.fields("CProvincia") & " " & AdoPaises.Recordset.fields("Descripcion_Rubro")
     Do While Not AdoPaises.Recordset.EOF
        CProvincia.AddItem AdoPaises.Recordset.fields("CProvincia") & " " & AdoPaises.Recordset.fields("Descripcion_Rubro")
        AdoPaises.Recordset.MoveNext
     Loop
  End If
  LlenarEmpresa DCListEmp
  RatonNormal
End Sub

Private Sub Form_Load()
  HayCnn = Get_WAN_IP
  CentrarForm CrearEmp
  ConectarAdodc AdoClave
  ConectarAdodc AdoPaises
  ConectarAdodc AdoEmpresa
  ConectarAdodc AdoBusqEmp
  ConectarAdodc AdoListEmp
  ConectarAdodc AdoListEmpCopy
  ConectarAdodc_MySQL AdoMySQLClave
End Sub

Private Sub OpcCDM_Click()
 If OpcCDM.value = 1 Then OpcCD.value = 0 Else OpcCD.value = 1
End Sub

Private Sub OpcCEM_Click()
  If OpcCEM.value = 1 Then OpcCE.value = 0 Else OpcCE.value = 1
End Sub

Private Sub OpcCIM_Click()
  If OpcCIM.value = 1 Then OpcCI.value = 0 Else OpcCI.value = 1
End Sub

Private Sub OpcCD_Click()
 If OpcCD.value = 1 Then OpcCDM.value = 0 Else OpcCDM.value = 1
End Sub

Private Sub OpcCE_Click()
  If OpcCE.value = 1 Then OpcCEM.value = 0 Else OpcCEM.value = 1
End Sub

Private Sub OpcCI_Click()
  If OpcCI.value = 1 Then OpcCIM.value = 0 Else OpcCIM.value = 1
End Sub

Private Sub OpcAmbiente_Click()
  'Metodo OffLine
  TxtWebRecepcion = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantesOffline?wsdl"
  TxtWebAutorizacion = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"

'' 'Metodo OnLine
''  TxtWebRecepcion = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantes?wsdl"
''  TxtWebAutorizacion = "https://celcer.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantes?wsdl"
End Sub

Private Sub OpcProduccion_Click()
 'Metodo OffLine
  TxtWebRecepcion = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantesOffline?wsdl"
  TxtWebAutorizacion = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl"
''
'' 'Metodo OnLine
''  TxtWebRecepcion = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/RecepcionComprobantes?wsdl"
''  TxtWebAutorizacion = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantes?wsdl"
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
 Select Case Button.key
   Case "Salir"
        Unload CrearEmp
        ListEmp.Show
   Case "Eliminar"
        Eliminar
   Case "Grabar"
        Grabar_Empresa
   Case "Produccion"
        Encera_Comprobantes_Electronicos
 End Select
End Sub

Private Sub TextAbreviatura_GotFocus()
  MarcarTexto TextAbreviatura
End Sub

Private Sub TextAbreviatura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextAbreviatura_LostFocus()
   TextoValido TextAbreviatura
   If TextAbreviatura = Ninguno Then TextAbreviatura = "Ninguna"
End Sub

Private Sub TextClave_GotFocus()
  MarcarTexto TextClave
End Sub

Private Sub TextClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDireccion_GotFocus()
  MarcarTexto TextDireccion
End Sub

Private Sub TextDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDireccion_LostFocus()
  TextoValido TextDireccion
End Sub

Private Sub TextEmpresa_GotFocus()
  MarcarTexto TextEmpresa
End Sub

Private Sub TextEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEmpresa_LostFocus()
  If TextEmpresa.Text = "" Then TextEmpresa.Text = "DEFAULT"
End Sub

Private Sub TextFAX_GotFocus()
  MarcarTexto TextFAX
End Sub

Private Sub TextFAX_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFAX_LostFocus()
  TextoValido TextFAX
End Sub

Private Sub TextGerente_GotFocus()
  MarcarTexto TextGerente
End Sub

Private Sub TextGerente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextGerente_LostFocus()
  TextoValido TextGerente
  TextGerente = ULCase(TextGerente)
End Sub

Private Sub TextLogoTipo_GotFocus()
  MarcarTexto TextLogoTipo
End Sub

Private Sub TextLogoTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextLogoTipo_LostFocus()
  TextoValido TextLogoTipo, False, True
  SSTab1.Tab = 0
  TextEmpresa.SetFocus
End Sub

Private Sub TextRUC_GotFocus()
  MarcarTexto TextRUC
End Sub

Private Sub TextRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRUC_LostFocus()
  If TextRUC = "" Then TextRUC = LimpiarRUC
  DigVerif = Digito_Verificador(TextRUC)
  CObligado.Text = TipoSRI.Obligado
  If Tipo_RUC_CI.Tipo_Beneficiario <> "R" Then MsgBox "RUC Incorrecto"
End Sub

Private Sub TextS_M_GotFocus()
  MarcarTexto TextS_M
End Sub

Private Sub TextS_M_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextS_M_LostFocus()
  TextoValido TextS_M, False
  If TextS_M = Ninguno Then TextS_M = "USD"
End Sub

Private Sub TextSubDir_GotFocus()
  MarcarTexto TextSubDir
End Sub

Private Sub TextSubDir_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextSubDir_LostFocus()
Dim NumEmpSubDir As Integer
  TextoValido TextSubDir, , True
  sSQL = "SELECT " & Full_Fields("Empresas") & " " _
       & "FROM Empresas " _
       & "WHERE Item <> '000' " _
       & "ORDER BY Item DESC "
  Select_Adodc AdoBusqEmp, sSQL
  With AdoBusqEmp.Recordset
    If TextSubDir = Ninguno Then
       NumEmpSubDir = 0
       If .RecordCount > 0 Then
          .MoveFirst
           NumEmpSubDir = Val(.fields("Item"))
       End If
       NumEmpSubDir = NumEmpSubDir + 1
       TextSubDir = "EMPRE" & Format$(NumEmpSubDir, "000")
    Else
      .MoveFirst
      .Find ("SubDir = '" & TextSubDir & "'")
       If Not .EOF Then
          If NumEmpresa <> .fields("Item") Then
             MsgBox "Este directorio ya existe seleccione otro"
             TextSubDir.SetFocus
          End If
       End If
    End If
  End With
End Sub

Private Sub TextTelefono1_GotFocus()
  MarcarTexto TextTelefono1
End Sub

Private Sub TextTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTelefono1_LostFocus()
  If TextTelefono1.Text = "" Then TextTelefono1.Text = LimpiarTelef
End Sub

Private Sub TextTelefono2_GotFocus()
  MarcarTexto TextTelefono2
End Sub

Private Sub TextTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTelefono2_LostFocus()
  If TextTelefono2.Text = "" Then TextTelefono2.Text = LimpiarTelef
End Sub

Public Sub LlenarEmpresa(NombreEmpresa As String)
Dim LogoTipo1 As String
  With AdoListEmp.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Empresa = '" & NombreEmpresa & "' ")
       If Not .EOF Then
          NumEmpresa = .fields("Item")
          TextRUC = .fields("RUC")
          TextEmpresa = .fields("Empresa")
          TextGerente = .fields("Gerente")
          TextLogoTipo = .fields("Logo_Tipo")
          TextTelefono1 = .fields("Telefono1")
          TextTelefono2 = .fields("Telefono2")
          TextFAX = .fields("FAX")
          TextS_M = .fields("S_M")
          TextDireccion = .fields("Direccion")
          TextSubDir = .fields("SubDir")
          'TextIVA = .Fields("IVA")
          TxtCodBanco = .fields("CodBanco")
          TxtContador = .fields("Contador")
          TxtRUCCont = .fields("RUC_Contador")
          TxtCI = .fields("CI_Representante")
          TxtRazonSocial = .fields("Razon_Social")
          TxtComercial = .fields("Nombre_Comercial")
          TxtEmail = .fields("Email")
          TxtEmailContador = .fields("Email_Contabilidad")
          TxtDecCosto = .fields("Dec_Costo")
          'TxtMeses = .Fields("Meses_Provision")
          TxtDecPVP = .fields("Dec_PVP")
          TxtDecCosto = .fields("Dec_Costo")
          TxtDecIVA = .fields("Dec_IVA")
          TxtDecCant = .fields("Dec_Cant")
          TxtSeguro = .fields("Seguro")
          TxtSeguro2 = .fields("Seguro2")
          TextAbreviatura = .fields("Abreviatura")
          TxtEstablecimientos = .fields("Establecimientos")
          TxtEmailConexion = .fields("Email_Conexion")
          TxtPasword = .fields("Email_Contraseña")
          TxtEmailConexionCE = .fields("Email_Conexion_CE")
          TxtPaswordCE = .fields("Email_Contraseña_CE")
          TxtLeyendaFA = .fields("LeyendaFA")
          TxtLeyendaFA1 = .fields("LeyendaFAT")
          TxtEmailRespaldo = .fields("Email_Respaldos")
          TxtEmailProcesos = .fields("Email_Procesos")
          TxtWebRecepcion = .fields("Web_SRI_Recepcion")
          TxtWebAutorizacion = .fields("Web_SRI_Autorizado")
          TxtCertificado = .fields("Ruta_Certificado")
          TxtPwdCertificado = .fields("Clave_Certificado")
          TxtContEsp = .fields("Codigo_Contribuyente_Especial")
          TxtTipoCarga = .fields("Tipo_Carga_Banco")
          CObligado.Text = .fields("Obligado_Conta")
          TxtRUCOperadora.Text = .fields("RUC_Operadora")
          
         'If .Fields("No_ATS") Then CheqNoATS.value = 1 Else CheqNoATS.value = 0
         'If .Fields("Sucursal") Then CheqSuc.value = 1 Else CheqSuc.value = 0
          If .fields("Det_Comp") Then CheqDetComp.value = 1 Else CheqDetComp.value = 0
          If .fields("Det_SubMod") Then CheqSubMod.value = 1 Else CheqSubMod.value = 0
          If .fields("Mod_Fact") Then CheqModFact.value = 1 Else CheqModFact.value = 0
          If .fields("Mod_PVP") Then CheqModPVP.value = 1 Else CheqModPVP.value = 0
          If .fields("Medio_Rol") Then CheqMedioRol.value = 1 Else CheqMedioRol.value = 0
          If .fields("Rol_2_Pagina") Then Cheq2Pag.value = 1 Else Cheq2Pag.value = 0
          If .fields("Registrar_IVA") Then CheqRegistrarIVA.value = 1 Else CheqRegistrarIVA.value = 0
          If .fields("Imp_Recibo_Caja") Then CheqRecibo.value = 1 Else CheqRecibo.value = 0
          If .fields("Num_CD") Then OpcCDM.value = 1 Else OpcCD.value = 1
          If .fields("Num_CE") Then OpcCEM.value = 1 Else OpcCE.value = 1
          If .fields("Num_CI") Then OpcCIM.value = 1 Else OpcCI.value = 1
          If .fields("Num_NC") Then OpcNCM.value = 1 Else OpcNC.value = 1
          If .fields("Num_ND") Then OpcNDM.value = 1 Else OpcND.value = 1
          If .fields("Email_CE_Copia") Then CheqConCopia.value = 1 Else CheqConCopia.value = 0
          If .fields("Ambiente") = "1" Then OpcAmbiente.value = True Else OpcProduccion.value = True
          
          If .fields("smtp_UseAuntentificacion") Then CheqAutentificacion.value = 1 Else CheqAutentificacion.value = 0
          If .fields("smtp_SSL") Then CheqSSL.value = 1 Else CheqSSL.value = 0
          If .fields("smtp_Secure") Then CheqSecure.value = 1 Else CheqSecure.value = 0
          TxtServidorSMTP.Text = .fields("smtp_Servidor")
          TxtPuerto.Text = .fields("smtp_Puerto")
          
         'Datos de la Empresa para averiguar si creamos un nueva
          TipoBenefCI = .fields("TD")
          'NomComer = .Fields("Nombre_Comercial")
          'NomEmp = .Fields("Empresa")
          'RUCEmp = .Fields("RUC")
          'NomGeren = .Fields("Gerente")
          'PrcIVA = .Fields("IVA")
          SSTab1.Tab = 0
          SSTab1.Caption = "Datos Principales (No. " & NumEmpresa & ")"
          LogoTipo1 = Ninguno
          If TextLogoTipo <> Ninguno Then
              LogoTipo1 = RutaSistema & "\LOGOS\"
              If Existe_File(LogoTipo1 & TextLogoTipo & ".gif") Then
                 LogoTipo1 = RutaSistema & "\LOGOS\" & TextLogoTipo & ".gif"
              ElseIf Existe_File(LogoTipo1 & TextLogoTipo & ".jpg") Then
                 LogoTipo1 = RutaSistema & "\LOGOS\" & TextLogoTipo & ".jpg"
              Else
                 LogoTipo1 = Ninguno
              End If
          End If
          If LogoTipo1 <> Ninguno Then
             Picture1.Picture = LoadPicture()
             Picture1.AutoRedraw = True
             Picture1.PaintPicture LoadPicture(LogoTipo1), 0.01, 0.01, 6400, 3200
          End If
          CCiudadS = UCaseStrg(.fields("Ciudad"))
          CNacion = .fields("CPais") & " " & .fields("Pais")
          CProvincia.Text = Ninguno
          For I = 0 To CProvincia.ListCount - 1
           If .fields("CProv") = MidStrg(CProvincia.List(I), 1, 2) Then
               CProvincia.Text = CProvincia.List(I)
           End If
          Next I
          For I = 0 To CNacion.ListCount - 1
           If .fields("CPais") = MidStrg(CNacion.List(I), 1, 3) Then
               CNacion.Text = CNacion.List(I)
           End If
          Next I
         'Verificamos si esta creada la empresa
          RUC = TextRUC
          Empresa = TextEmpresa
          RazonSocial = TxtRazonSocial
          NombreCiudad = CCiudadS
          NombreContador = TxtContador
          RUC_Contador = TxtRUCCont
          NombreGerente = TextGerente
          NLogoTipo = TextLogoTipo
          NMarcaAgua = "DISKCOVER"
          CadenaParcial = ""
         '|--=:******* CONECCON A MYSQL *******:=--|
          Datos_Iniciales_Entidad_SP_MySQL
         '|--=:******* --------.------- *******:=--|
       Else
          MsgBox "No existe esta Empresa"
       End If
   End If
  End With
End Sub

Private Sub TextUsuario_GotFocus()
  MarcarTexto TextUsuario
End Sub

Private Sub TextUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextUsuario_LostFocus()
  sSQL = "SELECT Codigo " _
       & "FROM Accesos " _
       & "WHERE UCaseStrg(Usuario) = '" & UCaseStrg(TextUsuario.Text) & "' "
  Select_Adodc AdoClave, sSQL
  If AdoClave.Recordset.RecordCount > 0 Then
     MsgBox "Este usuario ya existe, se procedera a actualizar la clave"
    TextClave.SetFocus
  End If
End Sub

Private Sub TxtCertificado_GotFocus()
  MarcarTexto TxtCertificado
End Sub

Private Sub TxtCertificado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCI_GotFocus()
  MarcarTexto TxtCI
End Sub

Private Sub TxtCI_LostFocus()
  If TxtCI = "" Then TxtCI = LimpiarCI
  DigVerif = Digito_Verificador(TxtCI)
  If (Tipo_RUC_CI.Tipo_Beneficiario = "P") Or (Tipo_RUC_CI.Tipo_Beneficiario = "C") Then
    'Correcto
  Else
     MsgBox "C.I. Incorrecta"
  End If
  TipoBenefCI = Tipo_RUC_CI.Tipo_Beneficiario
End Sub

Private Sub TxtCodBanco_GotFocus()
  MarcarTexto TxtCodBanco
End Sub

Private Sub TxtCodBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtComercial_GotFocus()
  MarcarTexto TxtComercial
End Sub

Private Sub TxtComercial_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtComercial_LostFocus()
  TextoValido TxtComercial, , True
  If Len(TxtComercial) <= 1 Then TxtComercial = TxtRazonSocial
End Sub

Private Sub TxtContador_GotFocus()
  MarcarTexto TxtContador
End Sub

Private Sub TxtContador_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtContEsp_GotFocus()
  MarcarTexto TxtContEsp
End Sub

Private Sub TxtContEsp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDecCant_GotFocus()
  MarcarTexto TxtDecCant
End Sub

Private Sub TxtDecCant_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDecCant_LostFocus()
  SSTab1.Tab = 1
  OpcAmbiente.SetFocus
End Sub

Private Sub TxtDecCosto_GotFocus()
  MarcarTexto TxtDecCosto
End Sub

Private Sub TxtDecCosto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDecIVA_GotFocus()
  MarcarTexto TxtDecIVA
End Sub

Private Sub TxtDecIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDecPVP_GotFocus()
  MarcarTexto TxtDecPVP
End Sub

Private Sub TxtDecPVP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  If Len(TxtEmail) <= 1 Then TxtEmail = "asistencia@diskcoversystem.com"
End Sub

Private Sub TxtEmailConexion_GotFocus()
  MarcarTexto TxtEmailConexion
End Sub

Private Sub TxtEmailConexion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailConexionCE_GotFocus()
  MarcarTexto TxtEmailConexionCE
End Sub

Private Sub TxtEmailConexionCE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailContador_GotFocus()
  MarcarTexto TxtEmailContador
End Sub

Private Sub TxtEmailContador_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailProcesos_GotFocus()
  MarcarTexto TxtEmailProcesos
End Sub

Private Sub TxtEmailProcesos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailProcesos_LostFocus()
   TxtEmailProcesos = LCase(TxtEmailProcesos)
End Sub

Private Sub TxtEmailRespaldo_GotFocus()
  MarcarTexto TxtEmailRespaldo
End Sub

Private Sub TxtEmailRespaldo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEstablecimientos_GotFocus()
  MarcarTexto TxtEstablecimientos
End Sub

Private Sub TxtEstablecimientos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEstablecimientos_LostFocus()
  TxtEstablecimientos = Format$(Val(TxtEstablecimientos), "000")
End Sub

Private Sub TxtLeyendaFA_GotFocus()
   MarcarTexto TxtLeyendaFA
End Sub

Private Sub TxtLeyendaFA_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyT Then TxtLeyendaFA = vLeyendaFA
End Sub

Private Sub TxtLeyendaFA1_GotFocus()
   MarcarTexto TxtLeyendaFA1
End Sub

Private Sub TxtLeyendaFA1_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyT Then TxtLeyendaFA1 = vLeyendaFA1
End Sub

Private Sub TxtLeyendaFA1_LostFocus()
  SSTab1.Tab = 0
  TextEmpresa.SetFocus
End Sub

Private Sub TxtNumPatronal_GotFocus()
  MarcarTexto TxtNumPatronal
End Sub

Private Sub TxtNumPatronal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPasword_GotFocus()
  MarcarTexto TxtPasword
End Sub

Private Sub TxtPasword_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPaswordCE_GotFocus()
  MarcarTexto TxtPaswordCE
End Sub

Private Sub TxtPaswordCE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPwdCertificado_GotFocus()
  MarcarTexto TxtPwdCertificado
End Sub

Private Sub TxtPwdCertificado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRazonSocial_GotFocus()
  MarcarTexto TxtRazonSocial
End Sub

Private Sub TxtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRazonSocial_LostFocus()
  TextoValido TxtRazonSocial, , True
End Sub

Private Sub TxtRUCCont_GotFocus()
  MarcarTexto TxtRUCCont
End Sub

Private Sub TxtRUCCont_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRUCCont_LostFocus()
  If TxtRUCCont = "" Then TxtRUCCont = LimpiarRUC
  DigVerif = Digito_Verificador(TxtRUCCont)
  If Tipo_RUC_CI.Tipo_Beneficiario <> "R" Then MsgBox "RUC Incorrecto"
End Sub

Private Sub OpcNC_Click()
  If OpcNC.value = 1 Then OpcNCM.value = 0 Else OpcNCM.value = 1
End Sub

Private Sub OpcNCM_Click()
  If OpcNCM.value = 1 Then OpcNC.value = 0 Else OpcNC.value = 1
End Sub

Private Sub OpcND_Click()
  If OpcND.value = 1 Then OpcNDM.value = 0 Else OpcNDM.value = 1
End Sub

Private Sub OpcNDM_Click()
  If OpcNDM.value = 1 Then OpcND.value = 0 Else OpcND.value = 1
End Sub

Private Sub TxtRUCOperadora_GotFocus()
   MarcarTexto TxtRUCOperadora
End Sub

Private Sub TxtRUCOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtRUCOperadora_LostFocus()
   TextoValido TxtRUCOperadora
   If Len(TxtRUCOperadora) = 13 Then
      DigVerif = Digito_Verificador(TxtRUCOperadora)
      If Tipo_RUC_CI.Tipo_Beneficiario <> "R" Then MsgBox "RUC de Operadora Incorrecto"
   End If
   SSTab1.Tab = 1
   CheqSubMod.SetFocus
End Sub

Private Sub TxtSeguro_GotFocus()
  MarcarTexto TxtSeguro
End Sub

Private Sub TxtSeguro_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSeguro2_GotFocus()
  MarcarTexto TxtSeguro
End Sub

Private Sub TxtSeguro2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTipoCarga_GotFocus()
  MarcarTexto TxtTipoCarga
End Sub

Private Sub TxtTipoCarga_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTipoCarga_LostFocus()
  TextoValido TxtTipoCarga, True, , 0
End Sub
