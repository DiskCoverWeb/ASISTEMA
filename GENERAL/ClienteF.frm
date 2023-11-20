VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FClientesFact 
   Caption         =   "ASIENTO DE MATRICULA"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Cliente"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modifica el Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar los Datos del Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "FactMult"
            Object.ToolTipText     =   "Asignar a Facturacion Multiple"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Facturacion"
            Object.ToolTipText     =   "Imprimir todas las Actas del Curso"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnFacMult"
            Object.ToolTipText     =   "Desabilitar Asignacion de Facturacion"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualiza Alumnos Recien Ingresados"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Matricula"
            Object.ToolTipText     =   "Registro Matricula"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "HojaMat"
            Object.ToolTipText     =   "Hoja de Matrícula"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Acta"
            Object.ToolTipText     =   "Acta de Matrícula"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificado"
            Object.ToolTipText     =   "Certificado de Matrícula"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Recalculo"
            Object.ToolTipText     =   "Certificado de Matrícula En Grupo"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Desactivar_Aliumnos"
            Object.ToolTipText     =   "Desactiva los alumnos de un curso "
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Carnet"
            Object.ToolTipText     =   "Imprime Carnet"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Certificado_Votar"
            Object.ToolTipText     =   "Certificado de Votacion"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Activar el Estudiante para asiento de matricula"
            Object.Tag             =   ""
            ImageIndex      =   24
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Renumerar_Actas"
            Object.ToolTipText     =   "Renumerar las actas de matriculas"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame3 
         Caption         =   "Fecha Impresion"
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
         Height          =   645
         Left            =   10815
         TabIndex        =   91
         Top             =   0
         Width           =   1695
         Begin MSMask.MaskEdBox MBFecha 
            Height          =   330
            Left            =   105
            TabIndex        =   92
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   210
            Width           =   1485
            _ExtentX        =   2619
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
   End
   Begin InetCtlsObjects.Inet URLinet 
      Left            =   11655
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " NOMBRE DEL ALUMNO(A)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   210
      TabIndex        =   1
      Top             =   840
      Width           =   10620
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "ClienteF.frx":0000
         DataSource      =   "AdoListCtas"
         Height          =   2325
         Left            =   105
         TabIndex        =   3
         ToolTipText     =   "Ctrl+B: Buscar datos en forma general"
         Top             =   525
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   4101
         _Version        =   393216
         Style           =   1
         ForeColor       =   8388608
         Text            =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtCodigoFoto 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0000000000"
         Top             =   2520
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo DCGrupo 
         Bindings        =   "ClienteF.frx":001A
         DataSource      =   "AdoGrupo"
         Height          =   315
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   16711680
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
      Begin VB.CommandButton Command1 
         Caption         =   "&S"
         Height          =   330
         Left            =   7980
         TabIndex        =   86
         Top             =   210
         Width           =   330
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FOTO ALUMNO(A)"
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
         Height          =   2220
         Left            =   8400
         TabIndex        =   85
         Top             =   210
         Width           =   2115
         Begin VB.Image ImgFoto 
            Height          =   1935
            Left            =   105
            Picture         =   "ClienteF.frx":0031
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1875
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7515
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   13256
      _Version        =   393216
      TabOrientation  =   3
      TabHeight       =   520
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MATRICULA"
      TabPicture(0)   =   "ClienteF.frx":1721E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label34"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label35"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label20"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label28"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label33"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label39"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label37"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label36"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label24"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label38"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label41"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label40"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label21"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "MBFechaN"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CCiudadS"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CProvincia"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "OpcF"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "OpcM"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtCI_RUC"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtApellidosS"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtDomicilio"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtTelefonoS"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtCIAlumno"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtProcedencia"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtNacionalidad"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "MBFechaM"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtNumero"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtFolio"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CGrupo"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtCodigo"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtEmail"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Command5"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtObservacion"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).ControlCount=   42
      TabCaption(1)   =   "FAMILIARES"
      TabPicture(1)   =   "ClienteF.frx":1723A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtTelefonoTrabajoM"
      Tab(1).Control(1)=   "TxtLugarTrabajoM"
      Tab(1).Control(2)=   "TxtProfesionM"
      Tab(1).Control(3)=   "TxtNacionalidadM"
      Tab(1).Control(4)=   "TxtMadre"
      Tab(1).Control(5)=   "TxtTelefonoTrabajoP"
      Tab(1).Control(6)=   "TxtLugarTrabajoP"
      Tab(1).Control(7)=   "TxtProfesionP"
      Tab(1).Control(8)=   "TxtNacionalidadP"
      Tab(1).Control(9)=   "TxtPadre"
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(11)=   "Label16"
      Tab(1).Control(12)=   "Label15"
      Tab(1).Control(13)=   "Label2"
      Tab(1).Control(14)=   "Label14"
      Tab(1).Control(15)=   "Label23"
      Tab(1).Control(16)=   "Label9"
      Tab(1).Control(17)=   "Label26"
      Tab(1).Control(18)=   "Label12"
      Tab(1).Control(19)=   "Label13"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "REPRESENTANTE"
      TabPicture(2)   =   "ClienteF.frx":17256
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtEmailR"
      Tab(2).Control(1)=   "TxtTelefonoR"
      Tab(2).Control(2)=   "TxtEmailRS"
      Tab(2).Control(3)=   "TxtLugarTrabajoR"
      Tab(2).Control(4)=   "TxtTelefonoRS"
      Tab(2).Control(5)=   "TxtCelular"
      Tab(2).Control(6)=   "TxtActividad"
      Tab(2).Control(7)=   "TxtProfesionR"
      Tab(2).Control(8)=   "TxtRepresentante"
      Tab(2).Control(9)=   "TxtCIR"
      Tab(2).Control(10)=   "TxtRazonSocial"
      Tab(2).Control(11)=   "TxtCedulaR"
      Tab(2).Control(12)=   "Label19"
      Tab(2).Control(13)=   "Label27"
      Tab(2).Control(14)=   "LblTD"
      Tab(2).Control(15)=   "Label31"
      Tab(2).Control(16)=   "Label42"
      Tab(2).Control(17)=   "Line1"
      Tab(2).Control(18)=   "Label45"
      Tab(2).Control(19)=   "Label44"
      Tab(2).Control(20)=   "Label3"
      Tab(2).Control(21)=   "Label29"
      Tab(2).Control(22)=   "Label32"
      Tab(2).Control(23)=   "Label30"
      Tab(2).Control(24)=   "Label43"
      Tab(2).Control(25)=   "LblTDR"
      Tab(2).Control(26)=   "Label25"
      Tab(2).Control(27)=   "Label17"
      Tab(2).ControlCount=   28
      Begin VB.TextBox TxtEmailR 
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
         Left            =   -72585
         MaxLength       =   60
         TabIndex        =   84
         Top             =   6930
         Width           =   7260
      End
      Begin VB.TextBox TxtTelefonoR 
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
         Left            =   -67545
         MaxLength       =   50
         TabIndex        =   66
         Top             =   6615
         Width           =   2220
      End
      Begin VB.TextBox TxtEmailRS 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -72585
         MaxLength       =   60
         TabIndex        =   70
         Top             =   4410
         Width           =   7260
      End
      Begin VB.TextBox TxtLugarTrabajoR 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   68
         Top             =   4095
         Width           =   7260
      End
      Begin VB.TextBox TxtTelefonoRS 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   97
         Top             =   3780
         Width           =   7260
      End
      Begin VB.TextBox TxtObservacion 
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
         Left            =   2205
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   6195
         Width           =   8520
      End
      Begin VB.TextBox TxtCelular 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   20
         TabIndex        =   82
         Top             =   6615
         Width           =   2745
      End
      Begin VB.TextBox TxtActividad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   30
         TabIndex        =   80
         Top             =   6300
         Width           =   7260
      End
      Begin VB.TextBox TxtProfesionR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   78
         Top             =   5985
         Width           =   7260
      End
      Begin VB.TextBox TxtRepresentante 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   76
         Top             =   5670
         Width           =   7260
      End
      Begin VB.TextBox TxtCIR 
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
         Left            =   -72585
         MaxLength       =   13
         TabIndex        =   73
         Top             =   5355
         Width           =   2745
      End
      Begin VB.TextBox TxtRazonSocial 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   64
         Top             =   3465
         Width           =   7260
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7035
         TabIndex        =   94
         Top             =   4305
         Width           =   330
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
         Left            =   2205
         MaxLength       =   60
         TabIndex        =   39
         Top             =   5880
         Width           =   8520
      End
      Begin VB.TextBox TxtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   9240
         MaxLength       =   10
         TabIndex        =   37
         Text            =   "0000000000"
         Top             =   5565
         Width           =   1485
      End
      Begin VB.ComboBox CGrupo 
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
         Left            =   960
         TabIndex        =   93
         Text            =   "Combo1"
         Top             =   4305
         Width           =   2055
      End
      Begin VB.TextBox TxtFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9555
         MaxLength       =   10
         TabIndex        =   20
         ToolTipText     =   "<Ctrl+F1> Matricula Diurna, <Ctrl+F2> Matricula Vespertina"
         Top             =   4305
         Width           =   1170
      End
      Begin VB.TextBox TxtNumero 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   18
         ToolTipText     =   "<Ctrl+F1> Matricula Diurna, <Ctrl+F2> Matricula Vespertina"
         Top             =   4305
         Width           =   1170
      End
      Begin MSMask.MaskEdBox MBFechaM 
         Height          =   330
         Left            =   4095
         TabIndex        =   16
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   4305
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
      Begin VB.TextBox TxtNacionalidad 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   28
         Top             =   4935
         Width           =   2535
      End
      Begin VB.TextBox TxtProcedencia 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   36
         Top             =   5565
         Width           =   7050
      End
      Begin VB.TextBox TxtCedulaR 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -72585
         MaxLength       =   13
         TabIndex        =   61
         Top             =   3150
         Width           =   2745
      End
      Begin VB.TextBox TxtCIAlumno 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "9999999999"
         Top             =   4620
         Width           =   2325
      End
      Begin VB.TextBox TxtTelefonoS 
         BeginProperty Font 
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
         MaxLength       =   20
         TabIndex        =   24
         Top             =   4620
         Width           =   1380
      End
      Begin VB.TextBox TxtDomicilio 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   34
         Top             =   5250
         Width           =   4635
      End
      Begin VB.TextBox TxtTelefonoTrabajoM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   59
         Top             =   6195
         Width           =   4635
      End
      Begin VB.TextBox TxtLugarTrabajoM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   57
         Top             =   5880
         Width           =   4635
      End
      Begin VB.TextBox TxtProfesionM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   55
         Top             =   5565
         Width           =   4635
      End
      Begin VB.TextBox TxtNacionalidadM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   53
         Top             =   5250
         Width           =   4635
      End
      Begin VB.TextBox TxtMadre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   51
         Top             =   4935
         Width           =   4635
      End
      Begin VB.TextBox TxtTelefonoTrabajoP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   49
         Top             =   4410
         Width           =   4635
      End
      Begin VB.TextBox TxtLugarTrabajoP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   47
         Top             =   4095
         Width           =   4635
      End
      Begin VB.TextBox TxtProfesionP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   45
         Top             =   3780
         Width           =   4635
      End
      Begin VB.TextBox TxtNacionalidadP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   43
         Top             =   3465
         Width           =   4635
      End
      Begin VB.TextBox TxtPadre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72585
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3150
         Width           =   4635
      End
      Begin VB.TextBox TxtApellidosS 
         BeginProperty Font 
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
         TabIndex        =   6
         Top             =   3360
         Width           =   7470
      End
      Begin VB.TextBox TxtCI_RUC 
         BeginProperty Font 
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
         MaxLength       =   13
         TabIndex        =   8
         ToolTipText     =   "<Ctrl+M> Codigo de Matrícula"
         Top             =   3360
         Width           =   1800
      End
      Begin VB.OptionButton OpcM 
         Caption         =   "Masculino"
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
         Height          =   225
         Left            =   9450
         TabIndex        =   9
         Top             =   3150
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OpcF 
         Caption         =   "Femenino"
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
         Height          =   225
         Left            =   9450
         TabIndex        =   10
         Top             =   3360
         Width           =   1275
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
         Left            =   6090
         TabIndex        =   30
         Text            =   "PICHINCHA"
         Top             =   4935
         Width           =   4635
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
         Left            =   2205
         TabIndex        =   32
         Top             =   5250
         Width           =   2535
      End
      Begin MSMask.MaskEdBox MBFechaN 
         Height          =   330
         Left            =   2205
         TabIndex        =   22
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO"
         BeginProperty Font 
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
         TabIndex        =   65
         Top             =   6615
         Width           =   2325
      End
      Begin VB.Label Label27 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CORREO FACTURA"
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
         Left            =   -74895
         TabIndex        =   69
         Top             =   4410
         Width           =   2325
      End
      Begin VB.Label LblTD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
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
         Left            =   -69855
         TabIndex        =   62
         Top             =   3150
         Width           =   330
      End
      Begin VB.Label Label31 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LUGAR TRABAJO"
         BeginProperty Font 
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
         TabIndex        =   67
         Top             =   4095
         Width           =   2325
      End
      Begin VB.Label Label42 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO"
         BeginProperty Font 
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
         TabIndex        =   98
         Top             =   3780
         Width           =   2325
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         X1              =   -64290
         X2              =   -74895
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label Label45 
         Caption         =   "DATOS DEL REPRESENTANTE DEL ESTUDIANTE:"
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
         Height          =   225
         Left            =   -74895
         TabIndex        =   71
         Top             =   5040
         Width           =   5055
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CORREO ELECTRONICO"
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
         Left            =   -74895
         TabIndex        =   83
         Top             =   6930
         Width           =   2325
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OBSERVACIONES"
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
         Height          =   330
         Left            =   105
         TabIndex        =   96
         Top             =   6195
         Width           =   2115
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CELULAR"
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
         Left            =   -74895
         TabIndex        =   81
         Top             =   6615
         Width           =   2325
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OCUPACION"
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
         Left            =   -74895
         TabIndex        =   79
         Top             =   6300
         Width           =   2325
      End
      Begin VB.Label Label32 
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
         Height          =   330
         Left            =   -74895
         TabIndex        =   77
         Top             =   5985
         Width           =   2325
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " REPRESENTANTE"
         BeginProperty Font 
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
         TabIndex        =   75
         Top             =   5670
         Width           =   2325
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CEDULA DE IDENTIDAD"
         BeginProperty Font 
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
         TabIndex        =   72
         Top             =   5355
         Width           =   2325
      End
      Begin VB.Label LblTDR 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
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
         TabIndex        =   74
         Top             =   5355
         Width           =   330
      End
      Begin VB.Label Label25 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RAZON SOCIAL REPRE."
         BeginProperty Font 
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
         TabIndex        =   63
         Top             =   3465
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CORREO DEL ESTU."
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
         Height          =   330
         Left            =   105
         TabIndex        =   38
         Top             =   5880
         Width           =   2115
      End
      Begin VB.Label Label41 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PAG. No."
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
         Height          =   330
         Left            =   8505
         TabIndex        =   19
         Top             =   4305
         Width           =   1065
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOMO No."
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
         Height          =   330
         Left            =   5355
         TabIndex        =   17
         Top             =   4305
         Width           =   1695
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Matricula:"
         BeginProperty Font 
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
         TabIndex        =   15
         Top             =   4305
         Width           =   1065
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CURSO"
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
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   4305
         Width           =   855
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BACHILLERATO EN "
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
         TabIndex        =   11
         Top             =   3675
         Width           =   6420
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GRADO/PARALELO"
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
         Left            =   6510
         TabIndex        =   12
         Top             =   3675
         Width           =   4215
      End
      Begin VB.Label Label37 
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
         Left            =   9240
         TabIndex        =   88
         Top             =   7035
         Width           =   1485
      End
      Begin VB.Label Label39 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DEUDA PENDIENTE"
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
         Left            =   6825
         TabIndex        =   89
         Top             =   7035
         Width           =   2430
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROCEDENCIA"
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
         Height          =   330
         Left            =   105
         TabIndex        =   35
         Top             =   5565
         Width           =   2115
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CEDULA DE IDENTIDAD"
         BeginProperty Font 
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
         TabIndex        =   60
         Top             =   3150
         Width           =   2325
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CEDULA DE IDENTIDAD"
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
         Height          =   330
         Left            =   6090
         TabIndex        =   25
         Top             =   4620
         Width           =   2325
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DOMICILIO"
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
         Height          =   330
         Left            =   4725
         TabIndex        =   33
         Top             =   5250
         Width           =   1380
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
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   105
         TabIndex        =   31
         Top             =   5250
         Width           =   2115
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
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   4725
         TabIndex        =   29
         Top             =   4935
         Width           =   1380
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
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   105
         TabIndex        =   27
         Top             =   4935
         Width           =   2115
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MATRICULA"
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
         Height          =   330
         Left            =   3465
         TabIndex        =   23
         Top             =   4620
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "NO EXISTE DATOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   87
         Top             =   7035
         Width           =   4215
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO TRABAJO"
         BeginProperty Font 
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
         TabIndex        =   58
         Top             =   6195
         Width           =   2325
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LUGAR TRABAJO"
         BeginProperty Font 
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
         TabIndex        =   56
         Top             =   5880
         Width           =   2325
      End
      Begin VB.Label Label15 
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
         Height          =   330
         Left            =   -74895
         TabIndex        =   54
         Top             =   5565
         Width           =   2325
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   -74895
         TabIndex        =   52
         Top             =   5250
         Width           =   2325
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE DE LA MADRE"
         BeginProperty Font 
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
         TabIndex        =   50
         Top             =   4935
         Width           =   2325
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO TRABAJO"
         BeginProperty Font 
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
         TabIndex        =   48
         Top             =   4410
         Width           =   2325
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LUGAR TRABAJO"
         BeginProperty Font 
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
         TabIndex        =   46
         Top             =   4095
         Width           =   2325
      End
      Begin VB.Label Label26 
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
         Height          =   330
         Left            =   -74895
         TabIndex        =   44
         Top             =   3780
         Width           =   2325
      End
      Begin VB.Label Label12 
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
         Height          =   330
         Left            =   -74895
         TabIndex        =   42
         Top             =   3465
         Width           =   2325
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE DEL PADRE"
         BeginProperty Font 
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
         TabIndex        =   40
         Top             =   3150
         Width           =   2325
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " APELLIDOS Y NOMBRES"
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
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   3150
         Width           =   7470
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CODIGO BANCO"
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
         Height          =   225
         Left            =   7560
         TabIndex        =   7
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA NACIMIENTO"
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
         Height          =   330
         Left            =   105
         TabIndex        =   21
         Top             =   4620
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SECCION"
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
         Left            =   6510
         TabIndex        =   13
         Top             =   3990
         Width           =   4215
      End
   End
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   840
      Top             =   2310
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
      Caption         =   "Tarjetas"
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   840
      Top             =   2625
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
      Caption         =   "ListCtas"
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
      Left            =   840
      Top             =   1365
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   840
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
      Caption         =   "Cuentas"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   840
      Top             =   2940
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
      BackColor       =   -2147483644
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
      Caption         =   "Creditos"
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
      Left            =   840
      Top             =   1050
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
   Begin MSAdodcLib.Adodc AdoEducativo 
      Height          =   330
      Left            =   840
      Top             =   1680
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
      Caption         =   "Educativo"
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
      Left            =   840
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
   Begin MSAdodcLib.Adodc AdoDireccion 
      Height          =   330
      Left            =   840
      Top             =   3570
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
      Caption         =   "Direccion"
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
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":17272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":1758C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":178A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":17BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":17EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":181F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":1850E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":18828
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":18B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":18E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":19176
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":19490
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":28222
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":2853C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":28856
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":28B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":28D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":2953C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":2977A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":29A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":29DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":29F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":2A2A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ClienteF.frx":2A5BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FClientesFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Archivo_Foto As String
Dim Cliente_Ant As String

Public Sub GrabarCliente()
  Si_No = False
  T = Normal
  FechaValida MBFechaN
  FechaValida MBFechaM
  TextoValido TxtRazonSocial, , True
  TextoValido TxtApellidosS, , True
  TextoValido TxtPadre, , True
  TextoValido TxtMadre, , True
  TextoValido TxtCIAlumno, , True
  TextoValido TxtTelefonoS, , True
  TextoValido TxtPadre, , True
  TextoValido TxtMadre, , True
  TextoValido TxtNacionalidad, , True
  TextoValido TxtNacionalidadP, , True
  TextoValido TxtNacionalidadM, , True
  TextoValido TxtProfesionP, , True
  TextoValido TxtProfesionM, , True
  TextoValido TxtLugarTrabajoP, , True
  TextoValido TxtLugarTrabajoR, , True
  TextoValido TxtLugarTrabajoM, , True
  TextoValido TxtTelefonoTrabajoP, , True
  TextoValido TxtTelefonoTrabajoM, , True
  TextoValido TxtTelefonoR, , True
  TextoValido TxtTelefonoRS, , True
  TextoValido TxtCedulaR, , True
  TextoValido TxtDomicilio, , True
  TextoValido TxtObservacion, , True
  TextoValido TxtActividad, , True
  
  Mensajes = "Esta seguro de Grabar datos de:" & vbCrLf _
           & TxtApellidosS & vbCrLf _
           & "del Codigo: " & TxtCodigo
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes And AdoListCtas.Recordset.RecordCount > 0 Then
     Codigo = TxtCodigo
     If Codigo <> Ninguno Then
        RatonReloj
        sSQL = "SELECT * " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & Codigo & "' "
        Select_Adodc AdoAux, sSQL
        If AdoAux.Recordset.RecordCount > 0 Then
            SetFields AdoAux, "T", Normal
            SetFields AdoAux, "Fecha", MBFechaM
            SetFields AdoAux, "Fecha_N", MBFechaN
            SetFields AdoAux, "Cliente", TxtApellidosS
            SetFields AdoAux, "Telefono", TxtTelefonoS
            SetFields AdoAux, "TelefonoT", TxtTelefonoRS
            SetFields AdoAux, "Celular", TxtCelular
            SetFields AdoAux, "Grupo", CGrupo
            SetFields AdoAux, "DirNumero", TxtNumero
            SetFields AdoAux, "No_Dep", 0
            SetFields AdoAux, "Ciudad", CCiudadS
            SetFields AdoAux, "Representante", TxtRazonSocial
            SetFields AdoAux, "Cedula", TxtCedulaR
            SetFields AdoAux, "DireccionT", TxtLugarTrabajoR
            SetFields AdoAux, "Est_Civil", "S"
            SetFields AdoAux, "Archivo_Foto", TxtCodigoFoto
            SetFields AdoAux, "Email", TxtEmailRS
            SetFields AdoAux, "Email2", TxtEmail
            If OpcM.value Then SetFields AdoAux, "Sexo", "M" Else SetFields AdoAux, "Sexo", "F"
            SetFields AdoAux, "Prov", SinEspaciosIzq(CProvincia)
            SetFields AdoAux, "CodigoU", CodigoUsuario
            SetUpdate AdoAux
        End If
        sSQL = "DELETE * " _
             & "FROM Clientes_Matriculas " _
             & "WHERE Codigo = '" & Codigo & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Item = '" & NumEmpresa & "' "
        Ejecutar_SQL_SP sSQL
       'Grabamos Datos de los Alumnos
        DigVerif = Digito_Verificador(URLInet, TxtCedulaR)
        SetAdoAddNew "Clientes_Matriculas"
        SetAdoFields "T", Normal
        SetAdoFields "Codigo", Codigo
        SetAdoFields "Nacionalidad", TxtNacionalidad
        SetAdoFields "Procedencia", TxtProcedencia
        SetAdoFields "Nombre_Padre", TxtPadre
        SetAdoFields "Nacionalidad_P", TxtNacionalidadP
        SetAdoFields "Profesion_P", TxtProfesionP
        SetAdoFields "Profesion_M", TxtProfesionM
        SetAdoFields "Profesion_R", TxtProfesionR
        SetAdoFields "Lugar_Trabajo_P", TxtLugarTrabajoP
        SetAdoFields "Lugar_Trabajo_M", TxtLugarTrabajoM
        SetAdoFields "Lugar_Trabajo_R", TxtLugarTrabajoR
        SetAdoFields "Telefono_Trabajo_P", TxtTelefonoTrabajoP
        SetAdoFields "Nombre_Madre", TxtMadre
        SetAdoFields "Nacionalidad_M", TxtNacionalidadM
        SetAdoFields "Telefono_Trabajo_M", TxtTelefonoTrabajoM
        SetAdoFields "Representante", TxtRazonSocial
        SetAdoFields "Representante_Alumno", TxtRepresentante
        SetAdoFields "Cedula_R", TxtCedulaR
        SetAdoFields "TD", LblTD.Caption
        SetAdoFields "Telefono_R", TxtTelefonoR
        SetAdoFields "Telefono_RS", TxtTelefonoRS
        SetAdoFields "Lugar_Nac", CCiudadS
        SetAdoFields "Matricula_No", TxtNumero
        SetAdoFields "Folio_No", TxtFolio
        SetAdoFields "Fecha_N", MBFechaN
        SetAdoFields "Fecha_M", MBFechaM
        SetAdoFields "Actividad_R", TxtActividad
        SetAdoFields "Domicilio", TxtDomicilio
        SetAdoFields "Telefono_D", TxtTelefonoS
        SetAdoFields "Grupo_No", CGrupo
        SetAdoFields "Observaciones", TxtObservacion
        SetAdoFields "CI", TxtCIAlumno
        SetAdoFields "Email_R", TxtEmailR
        SetAdoFields "CI_R", TxtCIR
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        TBeneficiario.Codigo = Codigo
        TBeneficiario.Direccion = TxtLugarTrabajoR
        TBeneficiario.Telefono1 = TxtTelefonoR
        TBeneficiario.Representante = TxtRazonSocial
        TBeneficiario.CI_RUC = TxtCedulaR
        RatonNormal
     Else
        MsgBox "No se puede grabar este Codigo"
     End If
  End If
  
  Actualizar_Notas_Curso "Clientes_Facturacion", "GrupoNo", Codigo
  Actualizar_Notas_Curso "Clientes_Matriculas", "Grupo_No", Codigo
  Actualizar_Notas_Curso "Trans_Actas", "CodE", Codigo
  Actualizar_Notas_Curso "Trans_Asistencia", "CodE", Codigo
  Actualizar_Notas_Curso "Trans_Notas", "CodE", Codigo
  Actualizar_Notas_Curso "Trans_Notas_Auxiliares", "CodE", Codigo
  Actualizar_Notas_Curso "Trans_Notas_Grado", "CodE", Codigo
  Actualizar_Notas_Curso "Trans_Promedios", "CodE", Codigo
  
  sSQL = "UPDATE Clientes_Matriculas " _
       & "SET Representante = '" & TxtRazonSocial & "' " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Cedula_R = '" & TxtCedulaR & "' "
  Ejecutar_SQL_SP sSQL
  
  Estudiante_DBF.codest = Codigo
  Estudiante_DBF.cedular = TxtCedulaR
  Estudiante_DBF.fonopaga = TxtTelefonoR
  Estudiante_DBF.pagador = TxtRazonSocial
  Estudiante_DBF.direcpaga = TxtLugarTrabajoR
 'Actualizar_Pagos
  ListarClientes CliFact
End Sub

Public Sub ListarClientes(Optional BuscarCliente As Boolean, Optional Beneficiario As String)
Dim TextosCliente As String
Dim CamposCliente As String
  RatonReloj
  TxtEmail = Ninguno
  TextosCliente = TxtApellidosS
  If TextosCliente = "" Then TextosCliente = Ninguno
 'Cadena1 = Tipo_Acceso_Educativo("", "Grupo")
  CamposCliente = "C.Fecha_N, C.Codigo, C.Cliente, C.TD, C.CI_RUC, C.Telefono, C.Celular, C.Direccion, C.Grupo, " _
                & "C.DirNumero, C.Ciudad, C.Email, C.Email2, C.Archivo_Foto, C.Sexo, C.Prov "
  If UCase$(Modulo) = "EDUCATIVO" Then
     sSQL = "SELECT " & CamposCliente & ",CM.Grupo_No " _
          & "FROM Clientes As C,Clientes_Matriculas As CM " _
          & "WHERE C.Codigo <> '.' " _
          & "AND CM.Periodo = '" & Periodo_Contable & "' " _
          & "AND CM.Item = '" & NumEmpresa & "' "
     If BuscarCliente Then
        sSQL = sSQL & "AND CM.Grupo_No = '" & SinEspaciosIzq(DCGrupo) & "' "
     Else
        If Len(Beneficiario) > 1 Then sSQL = sSQL & "AND C.Cliente LIKE '%" & Beneficiario & "%' "
     End If
     sSQL = sSQL _
          & "AND C.Codigo = CM.Codigo " _
          & "ORDER BY C.Cliente "
  Else
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Codigo <> '.' " _
          & "AND FA <> " & Val(adFalse) & " "
     If BuscarCliente Then
        sSQL = sSQL & "AND Grupo = '" & SinEspaciosIzq(DCGrupo) & "' "
     Else
        If Len(Beneficiario) > 1 Then sSQL = sSQL & "AND Cliente LIKE '%" & Beneficiario & "%' "
     End If
     sSQL = sSQL & "ORDER BY Cliente "
  End If
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Cliente"
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      'MsgBox .RecordCount
      .MoveFirst
      .Find ("Cliente Like '" & TextosCliente & "*' ")
       If Not .EOF Then
          DCCliente.Text = .Fields("Cliente")
          TxtEmail = .Fields("Email")
       Else
         .MoveFirst
       End If
       Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
   Else
       Frame1.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: 000000"
   End If
  End With
''' 'Lista de Alumnos Matriculados
'''  sSQL = "SELECT * " _
'''       & "FROM Clientes_Matriculas " _
'''       & "WHERE Codigo <> '.' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "ORDER BY Codigo "
'''  Select_Adodc AdoCuentas, sSQL
  RatonNormal
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  TxtCodigo = Ninguno
  TxtCodigoFoto = "SINFOTO"
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "' ")
       If Not .EOF Then
          MBFechaN = .Fields("Fecha_N")
          TxtCodigo = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          TxtApellidosS = .Fields("Cliente")
          DCCliente = .Fields("Cliente")
          TxtCI_RUC = .Fields("CI_RUC")
         'TxtRazonSocial = .Fields("Representante")
          TxtTelefonoS = .Fields("Telefono")
          TxtCelular = .Fields("Celular")
          DCDirS = .Fields("Direccion")
          CGrupo = .Fields("Grupo")
          TipoBenef = .Fields("TD")
          TxtNumero = .Fields("DirNumero")
          CCiudadS.Text = .Fields("Ciudad")
          TxtEmailRS = .Fields("Email")
          TxtEmail = .Fields("Email2")
          Archivo_Foto = .Fields("Archivo_Foto")
          TxtCodigoFoto = Archivo_Foto
          If .Fields("Sexo") = "M" Then OpcM.value = True Else OpcF.value = True
          Label6.Caption = "* CODIGO BANCO [" & TipoBenef & "]"
          For I = 0 To CProvincia.ListCount - 1
           If .Fields("Prov") = Mid$(CProvincia.List(I), 1, 2) Then
               CProvincia = CProvincia.List(I)
           End If
          Next I
          TxtApellidosS.Enabled = False
         'Listamos Datos del Alumno
          Listar_Alumnos
       Else
          MsgBox "No Existe"
       End If
   Else
     MsgBox "No Existe"
   End If
  End With
End Sub

Private Sub CCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CGrupo_GotFocus()
  MarcarTexto CGrupo
End Sub

Private Sub CGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CGrupo_LostFocus()
  If CGrupo.Text = Ninguno Then CGrupo.Text = NumEmpresa
End Sub

Private Sub Command1_Click()
  Unload FClientesFact
End Sub

Private Sub Command5_Click()
     sSQL = "SELECT MAX(Matricula_No) As NumMat " _
          & "FROM Clientes_Matriculas " _
          & "WHERE Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          If IsNull(.Fields("NumMat")) Then NumComp = 1 Else NumComp = Val(.Fields("NumMat")) + 1
          TxtNumero = Format$(NumComp, "000000")
      End If
     End With
End Sub

Private Sub CProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CProvincia_LostFocus()
  CCiudadS.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & CodigoPais & "' " _
       & "AND CProvincia = '" & SinEspaciosIzq(CProvincia) & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CCiudadS.AddItem AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NombreTabla(20) As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyB Then
     Cadena = InputBox("INGRESE NOMBRE" & vbCrLf & "DEL BENEFICIARIO:", "BUSQUEDA", "")
     ListarClientes False, Cadena
     DCCliente.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyW Then
     Mensajes = "Des-activar el Alumno(a)s del Plantel. " _
              & "Una vez que se procesa no se podra recuperar de la base"
     Titulo = "Pregunta de Desactivacion"
     If BoxMensaje = vbYes Then
        NombreTabla(0) = "Trans_Asistencia"
        NombreTabla(1) = "Trans_Actas"
        NombreTabla(2) = "Trans_Notas"
        NombreTabla(3) = "Trans_Notas_Grado"
        NombreTabla(4) = "Trans_Promedios"
        NombreTabla(5) = "Clientes_Matriculas"
        For I = 0 To 5
            sSQL = "UPDATE " & NombreTabla(I) & " " _
                 & "SET X = '.' " _
                 & "WHERE Item <> '.' "
            Ejecutar_SQL_SP sSQL
        Next I
        For I = 0 To 4
            If SQL_Server Then
               sSQL = "UPDATE " & NombreTabla(I) & " " _
                    & "SET X = 'X' " _
                    & "FROM " & NombreTabla(I) & " As T,Catalogo_Estudiantil As CE "
            Else
               sSQL = "UPDATE " & NombreTabla(I) & " As T,Catalogo_Estudiantil As CE " _
                    & "SET T.X = 'X' "
            End If
            sSQL = sSQL _
                 & "WHERE T.Item = '" & NumEmpresa & "' " _
                 & "AND T.Periodo = '" & Periodo_Contable & "' " _
                 & "AND CE.TC = 'P' " _
                 & "AND T.Periodo = CE.Periodo " _
                 & "AND T.Item = CE.Item " _
                 & "AND Mid$(T.CodE,1,7) = CE.CodigoE "
            Ejecutar_SQL_SP sSQL
        Next I
        If SQL_Server Then
           sSQL = "UPDATE Clientes_Matriculas " _
                & "SET X = 'X' " _
                & "FROM Clientes_Matriculas As T,Catalogo_Estudiantil As CE "
        Else
           sSQL = "UPDATE Clientes_Matriculas As T,Catalogo_Estudiantil As CE " _
                & "SET T.X = 'X' "
        End If
        sSQL = sSQL _
             & "WHERE T.Item = '" & NumEmpresa & "' " _
             & "AND T.Periodo = '" & Periodo_Contable & "' " _
             & "AND CE.TC = 'P' " _
             & "AND T.Periodo = CE.Periodo " _
             & "AND T.Item = CE.Item " _
             & "AND T.Grupo_No = CE.CodigoE "
        Ejecutar_SQL_SP sSQL
        For I = 0 To 5
            sSQL = "DELETE * " _
                 & "FROM " & NombreTabla(I) & " " _
                 & "WHERE X <> 'X' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Ejecutar_SQL_SP sSQL
        Next I
     End If
  End If
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente
  TipoDoc = "M"
End Sub

Private Sub DCGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupo_LostFocus()
  ListarClientes True
End Sub

Private Sub Form_Activate()
  MBFecha = FechaSistema
  FechaComp = MBFecha
  FClientesFact.Caption = "CREACION DEL CLIENTE"
  If Modulo = "TUTORIA" Then
     TBarCliente.Buttons("Imprimir").Enabled = False
     TBarCliente.Buttons("Modificar").Enabled = False
     TBarCliente.Buttons("Grabar").Enabled = False
     TBarCliente.Buttons("FactMult").Enabled = False
     TBarCliente.Buttons("Facturacion").Enabled = False
     TBarCliente.Buttons("UnFacMult").Enabled = False
     TBarCliente.Buttons("Actualizar").Enabled = False
     TBarCliente.Buttons("Matricula").Enabled = False
     TBarCliente.Buttons("Acta").Enabled = False
     TBarCliente.Buttons("HojaMat").Enabled = False
     TBarCliente.Buttons("Certificado").Enabled = False
     TBarCliente.Buttons("Recalculo").Enabled = False
     TBarCliente.Buttons("Carnet").Enabled = False
     TBarCliente.Buttons("Certificado_Votar").Enabled = False
     TBarCliente.Buttons("Desactivar_Aliumnos").Enabled = False
  End If
  Actualiza_Cursos
  Leer_Periodo_Lectivo
  TxtCodigoFoto.Text = "SINFOTO"
  Cadena1 = Tipo_Acceso_Educativo("", "CodigoE")
  
  sSQL = "SELECT CE.*,CC.Descripcion " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Cursos As CC " _
       & "WHERE CE.TC = 'P' " _
       & "AND CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.Item = CC.Item " _
       & "AND CE.Periodo = CC.Periodo " _
       & "AND CE.CodigoE = CC.Curso " _
       & Cadena1 _
       & "ORDER BY CC.Descripcion "
  Select_Adodc AdoEducativo, sSQL
  
  CGrupo.Clear
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND LEN(Curso)> 4 " _
       & "ORDER BY Curso "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Do While Not AdoAux.Recordset.EOF
        CGrupo.AddItem AdoAux.Recordset.Fields("Curso")
        AdoAux.Recordset.MoveNext
     Loop
     CGrupo.Text = CGrupo.List(0)
  End If
  
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND CPais = '" & CodigoPais & "' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CProvincia.AddItem AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CProvincia.AddItem "99 OTRO"
     CProvincia.Text = "99 OTRO"
  End If
  
  CCiudadS.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & CodigoPais & "' " _
       & "AND CProvincia = '" & CodigoProv & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoAux, sSQL
  
  If AdoAux.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CCiudadS.AddItem AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
  
 'Actualizando los Representantes
'''  If Periodo_Contable = Ninguno Then
'''        If SQL_Server Then
'''           sSQL = "UPDATE Clientes " _
'''                & "SET Representante = CM.Representante, Cedula = CM.Cedula_R," _
'''                & "Fecha_N = CM.Fecha_N, DireccionT = CM.Lugar_Trabajo_R " _
'''                & "FROM Clientes As C,Clientes_Matriculas As CM "
'''        Else
'''           sSQL = "UPDATE Clientes As C,Clientes_Matriculas As CM " _
'''                & "SET C.Representante = CM.Representante, C.Cedula = CM.Cedula_R," _
'''                & "C.Fecha_N = CM.Fecha_N, C.DireccionT = CM.Lugar_Trabajo_R "
'''        End If
'''        sSQL = sSQL _
'''             & "WHERE CM.Item = '" & NumEmpresa & "' " _
'''             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
'''             & "AND C.Codigo = CM.Codigo " _
'''             & "AND C.Representante <> CM.Representante "
'''        Ejecutar_SQL_SP sSQL
'''  End If
  If UCase$(Modulo) = "EDUCATIVO" Then
     sSQL = "SELECT (Curso & ' - ' & Descripcion) As Niveles " _
          & "FROM Catalogo_Cursos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND LEN(Curso)> 4 " _
          & "ORDER BY Curso,Descripcion "
  Else
    sSQL = "SELECT Grupo,Direccion,(Grupo & ' - ' & Direccion) As Niveles " _
         & "FROM Clientes " _
         & "WHERE Codigo <> '.' " _
         & "AND FA <> " & Val(adFalse) & " " _
         & "GROUP BY Grupo,Direccion " _
         & "ORDER BY Grupo,Direccion "
  End If
  SelectDB_Combo DCGrupo, AdoGrupo, sSQL, "Niveles"
  ListarClientes CliFact
  RatonNormal
  FClientesFact.WindowState = vbMaximized
  If Nuevo Then
     TxtApellidosS.Text = NombreCliente
     TxtCodigo = "Ninguno"
     CGrupo.Text = NumEmpresa
     TxtApellidosS.SetFocus
  Else
     ListarCuenta DCCliente.Text
  End If
  Select Case Modulo
    Case "NOTAS"
         Command2.Enabled = False
         Command3.Enabled = False
         Command4.Enabled = False
    Case "GERENCIA"
          Command2.Enabled = False
          Command3.Enabled = False
          Command4.Enabled = False
  End Select
End Sub

Private Sub Form_Deactivate()
  FClientesFact.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   'CentrarForm FClientesFact
   If Modulo = "GERENCIA" Then
      Command2.Enabled = False
      Command3.Enabled = False
      Command4.Enabled = False
   End If
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoEducativo
   ConectarAdodc AdoDireccion
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  FechaComp = MBFecha
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

Private Sub MBFechaN_GotFocus()
  MarcarTexto MBFechaN
End Sub

Private Sub MBFechaN_LostFocus()
  FechaValida MBFechaN
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 0 Then TxtCI_RUC.SetFocus
  If SSTab1.Tab = 1 Then TxtPadre.SetFocus
  If SSTab1.Tab = 2 Then TxtCIR.SetFocus
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
Dim Secundaria1 As Boolean
Dim Secundaria2 As Boolean
 FechaComp = MBFecha
'MsgBox Button.key
 Select Case Button.key
   Case "Salir"
        RatonNormal
        Unload FClientesFact
   Case "Imprimir"
        If FormatoLibreta = "BIMESTRES" Then
           Imprimir_Acta_Matricula DCCliente
           Imprimir_Hoja_Matricula DCCliente
        Else
           Imprimir_Acta_Matricula_Periodos DCCliente
           Imprimir_Hoja_Matricula_Periodos DCCliente
        End If
   Case "Modificar"
        TxtApellidosS.Enabled = True
        TxtRazonSocial.Enabled = True
        CCiudadS.SetFocus
   Case "Grabar"
        GrabarCliente
   Case "FactMult"
        Mensajes = "Asignar Facturacion " & TxtApellidosS.Text & "."
        Titulo = "Pregunta de Facturacion Multiple"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS.Text
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "Facturacion"
        Select Case FormatoLibreta
          Case "BIMESTRES", "TRIMESTRE1"
               Imprimir_Acta_Matricula DCCliente.Text
          Case Else
               Imprimir_Acta_Matricula_Periodos DCCliente.Text, SinEspaciosIzq(DCGrupo)
        End Select
'''        Mensajes = "Asignar a Facturacion " & TxtApellidosS.Text & "."
'''        Titulo = "Pregunta de Facturacion"
'''        If BoxMensaje = vbYes Then
'''           CodigoCliente = TxtCodigo
'''           NombreCliente = TxtApellidosS.Text
'''           sSQL = "UPDATE Clientes " _
'''                & "SET FA = " & Val(adTrue) & " " _
'''                & "WHERE Codigo = '" & CodigoCliente & "' "
'''           Ejecutar_SQL_SP sSQL
'''        End If
   Case "UnFacMult"
        Mensajes = "Des-activar a Facturacion a: " & TxtApellidosS.Text & "."
        Titulo = "Pregunta de Facturacion Multiple"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS.Text
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adFalse) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "Actualizar"
        ListarClientes CliFact
   Case "Matricula"
        Imprimir_Registro_Matricula DCCliente.Text
   Case "Acta"
        Select Case FormatoLibreta
          Case "BIMESTRES", "TRIMESTRE1"
               Imprimir_Acta_Matricula DCCliente.Text
          Case Else
               Imprimir_Acta_Matricula_Periodos DCCliente.Text
        End Select
   Case "HojaMat"
        Imprimir_Hoja_Matricula DCCliente.Text
   Case "Certificado"
        Imprimir_Certificado_Matricula DCCliente
   Case "Recalculo"
        Imprimir_Certificado_Matricula DCCliente, SinEspaciosIzq(DCGrupo)
   Case "Carnet"
        Carnet_Del_Alumno SinEspaciosIzq(DCGrupo)
   Case "Certificado_Votar"
        Certificado_Votacion_Del_Alumno SinEspaciosIzq(DCGrupo)
   Case "Desactivar_Aliumnos"
        Mensajes = "Des-activar el Alumno(a): " & TxtApellidosS.Text & " del Plantel. " _
                 & "Una vez que se procesa no se podra recuperar de la base"
        Titulo = "Pregunta de Desactivacion"
        If BoxMensaje = vbYes Then
           Codigo = TxtGrupo
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           sSQL = "DELETE * " _
                & "FROM Trans_Notas " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Notas_Grado " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Actas " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Trans_Promedios " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           sSQL = "DELETE * " _
                & "FROM Clientes_Matriculas " _
                & "WHERE Codigo = '" & CodigoCliente & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "Activar"
        If ClaveAdministrador Then
           RatonReloj
           Cadena = InputBox("INGRESE EL MOTIVO DE LA ACTIVACION", "ACTIVACION DE MATRICULA", "ORDEN SUPERIOR")
           Control_Procesos "E", "MATRICULAR POR: " & Cadena
           sSQL = "UPDATE Clientes_Matriculas " _
                & "SET Matricular = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & TxtCodigo & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Ejecutar_SQL_SP sSQL
           RatonNormal
           MsgBox "PROCESO REALIZADO CON EXITO"
        End If
   Case "Renumerar_Actas"
        If ClaveAdministrador Then
           RatonReloj
           sSQL = "SELECT * " _
                & "FROM Clientes " _
                & "WHERE FA <> " & Val(adFalse) & " " _
                & "AND Mid$(Grupo,1,2) = '1.' " _
                & "ORDER BY Grupo,Cliente "
           Select_Adodc AdoListCtas, sSQL
           With AdoListCtas.Recordset
            If .RecordCount > 0 Then
                Contador = 1
                Do While Not .EOF
                   FClientesFact.Caption = "Primaria: " & Format$(I / .RecordCount, "00%")
                   I = I + 1
                   Codigo1 = Val(Mid$(.Fields("Grupo"), 1, 1))
                  .Fields("DirNumero") = Codigo1 & Format$(Contador, "000")
                   Contador = Contador + 1
                  .MoveNext
                Loop
               .UpdateBatch
            End If
           End With
           Contador = 1
           sSQL = "SELECT * " _
                & "FROM Clientes " _
                & "WHERE FA <> " & Val(adFalse) & " " _
                & "AND Mid$(Grupo,1,2) = '2.' " _
                & "ORDER BY Grupo,Cliente "
           Select_Adodc AdoListCtas, sSQL
           With AdoListCtas.Recordset
            If .RecordCount > 0 Then
                Codigo1 = Trim$(Mid$(Anio_Lectivo, 3, 2))
                Do While Not .EOF
                   FClientesFact.Caption = "Secundaria: " & Format$(I / .RecordCount, "00%")
                   I = I + 1
                  .Fields("DirNumero") = Codigo1 & Format$(Contador, "000")
                   Contador = Contador + 1
                  .MoveNext
                Loop
               .UpdateBatch
            End If
           End With
           sSQL = "SELECT * " _
                & "FROM Clientes " _
                & "WHERE FA <> " & Val(adFalse) & " " _
                & "AND Mid$(Grupo,1,2) >= '3.' " _
                & "ORDER BY Grupo,Cliente "
           Select_Adodc AdoListCtas, sSQL
           With AdoListCtas.Recordset
            If .RecordCount > 0 Then
                Codigo1 = Trim$(Mid$(Anio_Lectivo, 3, 2))
                Do While Not .EOF
                   FClientesFact.Caption = "Bachillerato: " & Format$(I / .RecordCount, "00%")
                   I = I + 1
                  .Fields("DirNumero") = Codigo1 & Format$(Contador, "000")
                   Contador = Contador + 1
                  .MoveNext
                Loop
               .UpdateBatch
            End If
           End With
          'Lista de Alumnos Matriculados
           If SQL_Server Then
              sSQL = "UPDATE Clientes_Matriculas " _
                   & "SET Matricula_No = C.DirNumero,Folio_No = C.DirNumero " _
                   & "FROM Clientes_Matriculas As CM,Clientes As C "
           Else
              sSQL = "UPDATE Clientes_Matriculas As CM,Clientes As C " _
                   & "SET CM.Matricula_No = C.DirNumero,CM.Folio_No = C.DirNumero "
           End If
           sSQL = sSQL & "WHERE C.FA <> " & Val(adFalse) & " " _
                & "AND CM.Item = '" & NumEmpresa & "' " _
                & "AND CM.Periodo = '" & Periodo_Contable & "' " _
                & "AND ISNUMERIC(C.DirNumero) <> 0 " _
                & "AND CM.Codigo = C.Codigo "
           Ejecutar_SQL_SP sSQL
          'Presentamos el cursos actual
           ListarClientes True
           RatonNormal
           FClientesFact.Caption = "REGISTRO DE MATRICULAS"
           MsgBox "Proceso Terminado"
        End If
 End Select
 If Button.key <> "Salir" Then ListarClientes True
 RatonNormal
End Sub

Private Sub TxtActividad_GotFocus()
  MarcarTexto TxtActividad
End Sub

Private Sub TxtActividad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCedulaR_GotFocus()
  MarcarTexto TxtCedulaR
End Sub

Private Sub TxtCedulaR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCedulaR_LostFocus()
    If Len(TxtCedulaR) > 1 Then
       DigVerif = Digito_Verificador(URLInet, TxtCedulaR)
       If Tipo_RUC_CI.Tipo_Beneficiario = "P" Then
          Mensajes = "Este código es un Pasaporte"
          Titulo = "CONFIRMACION DE PASAPORTE"
          If BoxMensaje <> vbYes Then Tipo_RUC_CI.Tipo_Beneficiario = "O"
       End If
       If DigVerif = "-" Then
          MsgBox "RUC/CEDULA INCORRECTA"
          TxtCedulaR.SetFocus
       Else
          LblTD.Caption = Tipo_RUC_CI.Tipo_Beneficiario
          If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
             Label25.Caption = " R A Z O N    S O C I A L"
             Label7.Caption = " R.U.C."
          Else
             Label25.Caption = " APELLIDOS Y NOMBRES"
             Label7.Caption = " C.I./Pasaporte/Otros"
          End If
          LblTD.Caption = Tipo_RUC_CI.Tipo_Beneficiario
       End If
       sSQL = "SELECT CM.Codigo,CM.Representante,CM.Cedula_R,CM.Telefono_RS,C.Email " _
            & "FROM Clientes_Matriculas As CM, Clientes As C " _
            & "WHERE CM.Item = '" & NumEmpresa & "' " _
            & "AND CM.Periodo = '" & Periodo_Contable & "' " _
            & "AND CM.Cedula_R = '" & TxtCedulaR & "' " _
            & "AND CM.Codigo = C.Codigo "
       Select_Adodc AdoAux, sSQL
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            TxtRazonSocial = .Fields("Representante")
            TxtTelefonoRS = .Fields("Telefono_RS")
            TxtEmailRS = .Fields("Email")
        End If
       End With
    End If
End Sub

Private Sub TxtCI_RUC_GotFocus()
  MarcarTexto TxtCI_RUC
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyA Then
     sSQL = "UPDATE Clientes " _
          & "SET T = 'N' " _
          & "WHERE T = '.' "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "UPDATE Clientes_Matriculas " _
          & "SET T = 'N' " _
          & "WHERE T = '.' "
     Ejecutar_SQL_SP sSQL

     sSQL = "UPDATE Clientes_Matriculas " _
          & "SET Bachiller = '" & TxtBachiller & "', " _
          & "Especialidad = '" & TxtEspecialidad & "', " _
          & "Ciclo = '" & TxtCiclo & "', " _
          & "Seccion = '" & CSeccion & "', " _
          & "Nivel = '" & DCDirS & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Grupo_No = '" & TxtGrupo & "' "
     Ejecutar_SQL_SP sSQL
     MsgBox "Actualización realizado con éxito"
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
  TextoValido TxtApellidosS, , True
  With AdoListCtas.Recordset
   If .RecordCount > 0 And TxtApellidosS.Text <> Ninguno Then
       RatonReloj
      .MoveFirst
      .Find ("Cliente Like '" & TxtApellidosS & "' ")
       RatonNormal
       If Not .EOF Then
          MsgBox "El Cliente " & TxtApellidosS _
               & ", ya existe, está asignado a " & vbCrLf & vbCrLf _
               & .Fields("Cliente") & vbCrLf & vbCrLf _
               & "Codigo: " & .Fields("CI_RUC")
          DCCliente.SetFocus
       End If
   End If
  End With
End Sub

Private Sub TxtCelular_GotFocus()
  MarcarTexto TxtCelular
End Sub

Private Sub TxtCelular_LostFocus()
  TextoValido TxtCelular, , True
  TxtCelular.Text = Format$(Val(TxtCelular.Text), "0000000000")
End Sub

Private Sub TxtCI_RUC_LostFocus()
  TextoValido TxtCI_RUC, , True
  With AdoListCtas.Recordset
   If .RecordCount > 0 And TxtCI_RUC.Text <> Ninguno Then
       RatonReloj
      .MoveFirst
      .Find ("CI_RUC Like '" & TxtCI_RUC.Text & "' ")
       RatonNormal
       If Not .EOF Then
          If .Fields("Cliente") <> TxtApellidosS.Text Then
              MsgBox "Este Código, está asignado a " & vbCrLf & vbCrLf & .Fields("Cliente")
              TxtCI_RUC.SetFocus
          Else
              TipoBenef = .Fields("TD")
          End If
       Else
          DigVerif = Digito_Verificador(URLInet, TxtCI_RUC.Text)
          Caracter = Mid$(TxtCI_RUC.Text, 10, 1)
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "C", "D", "R"
                 If DigVerif <> Caracter Then
                    MsgBox "Codigo Incorrecto"
                    TxtCI_RUC.SetFocus
                 End If
          End Select
       End If
   End If
  End With
  Label4.Caption = "* APELLIDOS Y NOMBRES"
End Sub

Private Sub TxtCIAlumno_GotFocus()
  MarcarTexto TxtCIAlumno
End Sub

Private Sub TxtCIAlumno_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCIR_GotFocus()
  MarcarTexto TxtCIR
End Sub

Private Sub TxtCIR_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtCIR_LostFocus()
  If Len(TxtCIR) > 1 Then
     DigVerif = Digito_Verificador(URLInet, TxtCIR)
     If Tipo_RUC_CI.Tipo_Beneficiario = "P" Then
        Mensajes = "Este código es un Pasaporte"
        Titulo = "CONFIRMACION DE PASAPORTE"
        If BoxMensaje <> vbYes Then Tipo_RUC_CI.Tipo_Beneficiario = "O"
     End If
     If DigVerif = "-" Then
        MsgBox "RUC/CEDULA INCORRECTA"
        TxtCIR.SetFocus
     Else
        LblTDR.Caption = Tipo_RUC_CI.Tipo_Beneficiario
        If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
           MsgBox "SOLO SE ADMITE CEDULA O PASAPORTE"
           TxtCIR.SetFocus
        Else
           Label25.Caption = " APELLIDOS Y NOMBRES"
           Label7.Caption = " C.I./Pasaporte/Otros"
           sSQL = "SELECT CM.Codigo,CM.Representante_Alumno,CM.CI_R,Email_R,C.Celular,CM.Telefono_R,CM.Email_R " _
                & "FROM Clientes_Matriculas As CM, Clientes As C " _
                & "WHERE CM.Item = '" & NumEmpresa & "' " _
                & "AND CM.Periodo = '" & Periodo_Contable & "' " _
                & "AND CM.CI_R = '" & TxtCIR & "' " _
                & "AND CM.Codigo = C.Codigo "
           Select_Adodc AdoAux, sSQL
           With AdoAux.Recordset
            If .RecordCount > 0 Then
                TxtRepresentante = .Fields("Representante_Alumno")
                TxtCelular = .Fields("Celular")
                TxtEmailR = .Fields("Email_R")
                TxtTelefonoR = .Fields("Telefono_R")
            End If
           End With
           LblTDR.Caption = Tipo_RUC_CI.Tipo_Beneficiario
        End If
     End If
  End If
End Sub

Private Sub TxtDomicilio_GotFocus()
  MarcarTexto TxtDomicilio
End Sub

Private Sub TxtDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  TextoValido TxtEmail
End Sub

Private Sub TxtEmailR_GotFocus()
  MarcarTexto TxtEmailR
End Sub

Private Sub TxtEmailR_LostFocus()
  TextoValido TxtEmailR
End Sub

Private Sub TxtEmailR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmailRS_GotFocus()
  MarcarTexto TxtEmailRS
End Sub

Private Sub TxtEmailRS_LostFocus()
  TextoValido TxtEmailRS
End Sub

Private Sub TxtEmailRS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabajoM_GotFocus()
   MarcarTexto TxtLugarTrabajoM
End Sub

Private Sub TxtLugarTrabajoM_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabajoP_GotFocus()
   MarcarTexto TxtLugarTrabajoP
End Sub

Private Sub TxtLugarTrabajoP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabajoR_GotFocus()
  MarcarTexto TxtLugarTrabajoR
End Sub

Private Sub TxtLugarTrabajoR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMadre_GotFocus()
   MarcarTexto TxtMadre
End Sub

Private Sub TxtMadre_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNacionalidad_GotFocus()
  MarcarTexto TxtNacionalidad
End Sub

Private Sub TxtNacionalidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNacionalidad_LostFocus()
  TextoValido TxtNacionalidad, , True
End Sub

Private Sub TxtNacionalidadM_GotFocus()
   MarcarTexto TxtNacionalidadM
End Sub

Private Sub TxtNacionalidadM_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNacionalidadP_GotFocus()
   MarcarTexto TxtNacionalidadP
End Sub

Private Sub TxtNacionalidadP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumero_GotFocus()
  MarcarTexto TxtNumero
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  NumComp = Val(TxtNumero)
  If CtrlDown And KeyCode = vbKeyF1 Then
     NumComp = ReadSetDataNum("Diario", True, False)
     sSQL = "SELECT MAX(Matricula_No) As NumMat " _
          & "FROM Clientes_Matriculas " _
          & "WHERE Seccion = '" & CSeccion.Text & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          If IsNull(.Fields("NumMat")) Then NumComp = 1 Else NumComp = Val(.Fields("NumMat")) + 1
          TxtNumero = Format$(NumComp, "000000")
      End If
     End With
  End If
  If CtrlDown And KeyCode = vbKeyM Then TxtNumero = Numero_De_Matricula(CGrupo)
End Sub

Private Sub TxtNumero_LostFocus()
  NumComp = Val(TxtNumero)
  If NumComp <= 0 Then NumComp = 1
  TxtNumero = Format$(NumComp, "000000")
End Sub

Private Sub TxtObservacion_GotFocus()
  MarcarTexto TxtObservacion
End Sub

Private Sub TxtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPadre_GotFocus()
   MarcarTexto TxtPadre
End Sub

Private Sub TxtPadre_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtProcedencia_GotFocus()
  MarcarTexto TxtProcedencia
End Sub

Private Sub TxtProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtProfesionM_GotFocus()
  MarcarTexto TxtProfesionM
End Sub

Private Sub TxtProfesionM_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtProfesionP_GotFocus()
  MarcarTexto TxtProfesionP
End Sub

Private Sub TxtProfesionP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtProfesionR_GotFocus()
  MarcarTexto TxtProfesionR
End Sub

Private Sub TxtProfesionR_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub TxtRepresentante_GotFocus()
MarcarTexto TxtRepresentante
End Sub

Private Sub TxtTelefonoR_GotFocus()
  MarcarTexto TxtTelefonoR
End Sub

Private Sub TxtTelefonoR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoS_GotFocus()
  MarcarTexto TxtTelefonoS
End Sub

Private Sub TxtTelefonoS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoS_LostFocus()
  TextoValido TxtTelefonoS, , True
  TxtTelefonoS.Text = Format$(Val(TxtTelefonoS.Text), "000000000")
End Sub

Public Sub Listar_Alumnos()
Dim Numero_Fact As Integer
    Fechas_Balances "Deuda Pendiente", MBFecha, MBFecha
    Label18.Caption = ""
    TxtProcedencia = Ninguno
    TxtPadre = Ninguno
    TxtMadre = Ninguno
    TxtProfesionP = Ninguno
    TxtProfesionM = Ninguno
    TxtProfesionR = Ninguno
    TxtLugarTrabajoP = Ninguno
    TxtLugarTrabajoM = Ninguno
    TxtLugarTrabajoR = Ninguno
    TxtNacionalidad = Ninguno
    TxtNacionalidadP = Ninguno
    TxtNacionalidadM = Ninguno
    TxtTelefonoTrabajoP = Ninguno
    TxtTelefonoTrabajoM = Ninguno
    TxtTelefonoR = Ninguno
    TxtTelefonoRS = Ninguno
    TxtCedulaR = Ninguno
    TxtActividad = Ninguno
    TxtNumero = "0"
    TxtFolio = "0"
    TxtCIAlumno = Ninguno
    TxtDomicilio = Ninguno
    TxtObservacion = Ninguno
    TxtRazonSocial = Ninguno
    TxtRepresentante = Ninguno
    TxtEmailR = Ninguno
    MBFechaN.Text = FechaSistema
    Numero_Fact = 0
    Label37.Caption = "0.00"
    Label39.Caption = "DEUDA PENDIENTE"
   'Label18.Caption = "No existe Datos del Alumno(a)"
    Si_No = False
    sSQL = "SELECT * " _
         & "FROM Clientes_Matriculas " _
         & "WHERE Codigo = '" & CodigoCliente & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' "
    Select_Adodc AdoCuentas, sSQL
    
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
       'MsgBox .RecordCount
       ImgFoto.Picture = LoadPicture()
       'MsgBox CodigoCliente   '101299557
       RutaDestino = RutaSistema & "\FOTOS\" & Archivo_Foto & ".JPG"
       If Dir(RutaDestino) <> "" Then
          ImgFoto.Picture = LoadPicture(RutaDestino)
         'MsgBox RutaDestino
       Else
          RutaDestino = RutaSistema & "\FOTOS\" & Archivo_Foto & ".GIF"
          If Dir(RutaDestino) <> "" Then ImgFoto.Picture = LoadPicture(RutaDestino)
       End If
      'Listamo los Datos de los Alumnos
       If Not .EOF Then
         'Si existe llenamos Datos
          Label18.Caption = ""
          Label36.Caption = Leer_Datos_del_Curso(.Fields("Grupo_No"))
          If Len(Dato_Curso.Especialidad) > 1 Then Label36.Caption = Label36.Caption & vbCrLf & "ESPECIALIDAD: " & Dato_Curso.Especialidad & " "
          
          Label1.Caption = "SECCIÓN: " & Dato_Curso.Seccion
          Label7.Caption = Dato_Curso.Descripcion
          TxtProcedencia = .Fields("Procedencia")
          TxtPadre.Text = .Fields("Nombre_Padre")
          TxtMadre.Text = .Fields("Nombre_Madre")
          TxtProfesionP = .Fields("Profesion_P")
          TxtProfesionM = .Fields("Profesion_M")
          TxtProfesionR = .Fields("Profesion_R")
          TxtLugarTrabajoP = .Fields("Lugar_Trabajo_P")
          TxtLugarTrabajoM = .Fields("Lugar_Trabajo_M")
          TxtLugarTrabajoR = .Fields("Lugar_Trabajo_R")
          TxtNacionalidad = .Fields("Nacionalidad")
          TxtNacionalidadP = .Fields("Nacionalidad_P")
          TxtNacionalidadM = .Fields("Nacionalidad_M")
          TxtTelefonoTrabajoP = .Fields("Telefono_Trabajo_P")
          TxtTelefonoTrabajoM = .Fields("Telefono_Trabajo_M")
          TxtTelefonoR = .Fields("Telefono_R")
          TxtTelefonoRS = .Fields("Telefono_RS")
          TxtCedulaR = .Fields("Cedula_R")
          LblTD.Caption = .Fields("TD")
          TxtActividad = .Fields("Actividad_R")
          TxtCIAlumno = .Fields("CI")
          TxtDomicilio = .Fields("Domicilio")
          TxtObservacion = .Fields("Observaciones")
          TxtRazonSocial = .Fields("Representante")
          TxtRepresentante = .Fields("Representante_Alumno")
          TxtEmailR = .Fields("Email_R")
          MBFechaM = .Fields("Fecha_M")
          MBFechaN = .Fields("Fecha_N")
          CCiudadS = .Fields("Lugar_Nac")
          TxtNumero = .Fields("Matricula_No")
          TxtFolio = .Fields("Folio_No")
          CGrupo = .Fields("Grupo_No")
          Si_No = .Fields("Matricular")
          TxtCIR = .Fields("CI_R")
         'Averiguamos el numero de matricula
          If Val(TxtNumero) = 0 Then
             sSQL = "SELECT MAX(Matricula_No) As Matricula_Max " _
                  & "FROM Clientes_Matriculas " _
                  & "WHERE Periodo = '" & Periodo_Contable & "' " _
                  & "AND Item = '" & NumEmpresa & "' "
             If Mid$(.Fields("Grupo_No"), 1, 1) = "1" Then
                sSQL = sSQL & "AND Mid$(Grupo_No,1,1) = '1' "
             Else
                sSQL = sSQL & "AND Mid$(Grupo_No,1,1) <> '1' "
             End If
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then
                TxtNumero = AdoAux.Recordset.Fields("Matricula_Max")
             End If
             sSQL = "SELECT MAX(Folio_No) As Folio_Max " _
                  & "FROM Clientes_Matriculas " _
                  & "WHERE Periodo = '" & Periodo_Contable & "' " _
                  & "AND Item = '" & NumEmpresa & "' "
             If Mid$(.Fields("Grupo_No"), 1, 1) = "1" Then
                sSQL = sSQL & "AND Mid$(Grupo_No,1,1) = '1' "
             Else
                sSQL = sSQL & "AND Mid$(Grupo_No,1,1) <> '1' "
             End If
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then
                TxtFolio = AdoAux.Recordset.Fields("Folio_Max")
             End If
             TxtNumero = Val(TxtNumero) + 1
             If EsImpar(Val(TxtNumero)) Then TxtFolio = Val(TxtFolio) + 1
          End If
          Label37.Caption = "0.00"
          sSQL = "SELECT CodigoC, SUM(Saldo_MN) As Deuda_Pendiente, COUNT(CodigoC) As CFacturas " _
               & "FROM Facturas " _
               & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND CodigoC = '" & TxtCodigo & "' " _
               & "AND Saldo_MN > 0 " _
               & "AND T <> 'A' " _
               & "GROUP BY CodigoC "
          Select_Adodc AdoAux, sSQL
          Codigos = ""
          If AdoAux.Recordset.RecordCount > 0 Then
             Numero_Fact = AdoAux.Recordset.Fields("CFacturas")
             Label37.Caption = Format$(AdoAux.Recordset.Fields("Deuda_Pendiente"), "#,##0.00")
             Label39.Caption = "DEUDA PENDIENTE (" & Numero_Fact & ")"
          End If
       End If
   Else
      MsgBox "Beneficiario no Asignado"
   End If
  End With
  If Numero_Fact > 1 Then
     TBarCliente.Buttons.Item(4).Enabled = False
     TBarCliente.Buttons.Item(12).Enabled = False
     If Si_No Then
        TBarCliente.Buttons.Item(4).Enabled = True
        TBarCliente.Buttons.Item(12).Enabled = True
     End If
     MsgBox "N O   D E B E   M A T R I C U L A R" & vbCrLf & vbCrLf _
          & "E S T E   A L U M N O(A) " & vbCrLf & vbCrLf _
          & "P O R Q U E   T I E N E   U N A" & vbCrLf & vbCrLf _
          & "D E U D A   P E N D I E N T E   D E" & vbCrLf & vbCrLf _
          & Numero_Fact & "  F A C T U R A S" & vbCrLf & vbCrLf _
          & "Q U E   S U M A N   USD " & Label37.Caption
  Else
     TBarCliente.Buttons.Item(4).Enabled = True
     TBarCliente.Buttons.Item(12).Enabled = True
  End If
  If Modulo = "TUTORIA" Then
     TBarCliente.Buttons.Item(4).Enabled = False
     TBarCliente.Buttons.Item(12).Enabled = False
  End If
End Sub

Private Sub TxtTelefonoTrabajoM_GotFocus()
  MarcarTexto TxtTelefonoTrabajoM
End Sub

Private Sub TxtTelefonoTrabajoM_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoTrabajoP_GotFocus()
  MarcarTexto TxtTelefonoTrabajoP
End Sub

Private Sub TxtTelefonoTrabajoP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Carnet_Del_Alumno(Curso As String)
Dim PosX As Single
Dim PosY As Single
Dim SizeLetra As Integer
Dim No_Carnet As Integer
Dim FotoCarnet As String
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
PosX = 0: PosY = 0
SizeLetra = 9
Pagina = 1
No_Carnet = 1
'Iniciamos la impresion
Printer.FontBold = False
sSQL = "SELECT Sexo,Cliente,CI_RUC,Direccion,Grupo,Archivo_Foto " _
       & "FROM Clientes " _
       & "WHERE Grupo = '" & Curso & "' " _
       & "ORDER BY Grupo,Cliente,Sexo DESC "
Select_Adodc AdoAux, sSQL
DataAnchoCampos InicioX, AdoAux, SizeLetra, TipoArialNarrow, Orientacion_Pagina, True
With AdoAux.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
       RutaDestino = RutaSistema & "\FOTOS\" & .Fields("Archivo_Foto") & ".JPG"
       If Dir(RutaDestino) <> "" Then
          FotoCarnet = RutaDestino
       Else
          RutaDestino = RutaSistema & "\FOTOS\" & .Fields("Archivo_Foto") & ".GIF"
          If Dir(RutaDestino) <> "" Then FotoCarnet = RutaDestino
       End If
       PrinterPaint FotoCarnet, PosX + 0.55, PosY + 1.55, 2, 1.9
       RutaDestino = RutaSistema & "\LOGOS\MINISEDU.jpg"
       PrinterPaint RutaDestino, PosX + 0.2, PosY + 0.1, 1.8, 1.2
       PrinterPaint LogoTipo, PosX + 6.5, PosY + 0.2, 2, 1
       
       Printer.Line (PosX + 0.5, PosY + 1.5)-(PosX + 2.6, PosY + 3.5), QBColor(0), B
       Printer.Line (PosX + 0.01, PosY + 0.01)-(PosX + 8.5, PosY + 5.5), QBColor(0), B
       
       Printer.FillColor = QBColor(Rojo)
       Printer.FontBold = True
       Printer.FontSize = 8
       PrinterTexto PosX + 2.1, PosY + 0.1, UCase$(Institucion1)
       Printer.FontSize = 6
       PrinterTexto PosX + 2.1, PosY + 0.5, Institucion2
       Printer.FillColor = QBColor(0)
       Printer.FontSize = 6
       Printer.FontItalic = True
       Printer.FontBold = False
       PrinterTexto PosX + 3, PosY + 0.9, Direccion
       PrinterTexto PosX + 3, PosY + 1.2, "Teléfono: " & Telefono1
       PrinterTexto PosX + 3, PosY + 1.4, UCase$(NombreCiudad) & " - MANABI"
       Printer.FontItalic = False
       Printer.FontBold = True
       Printer.FontSize = 8
       PrinterFields PosX + 2.7, PosY + 2, .Fields("Cliente")
       PrinterFields PosX + 2.7, PosY + 2.8, .Fields("Direccion")
       PrinterFields PosX + 0.5, PosY + 3.6, .Fields("CI_RUC")
       PrinterTexto PosX + 2.7, PosY + 3.6, Anio_Lectivo
       Printer.FontBold = False
       Printer.FontSize = 7
       PrinterTexto PosX + 2.7, PosY + 2.3, "ESTUDIANTE"
       PrinterTexto PosX + 2.7, PosY + 3.1, "CURSO"
       PrinterTexto PosX + 0.5, PosY + 3.9, "CÓDIGO"
       PrinterTexto PosX + 2.7, PosY + 3.9, "AÑO LECTIVO"
       
       PrinterTexto PosX + 0.5, PosY + 4.7, Rector
       PrinterTexto PosX + 5, PosY + 4.7, "Augusto Espinosa"
       PrinterTexto PosX + 0.5, PosY + 5, TextoRector
       PrinterTexto PosX + 5, PosY + 5, "Ministro de Educación"
       No_Carnet = No_Carnet + 1
       PosX = PosX + 8.5
       If No_Carnet > 2 Then
          PosX = 0
          PosY = PosY + 5.5
          No_Carnet = 1
       End If
       If PosY >= 25 Then
          Printer.NewPage
          PosY = 0
       End If
      .MoveNext
     Loop
End If
End With
RatonNormal
MensajeEncabData = "    "
Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Certificado_Votacion_Del_Alumno(Curso As String)
Dim PosX As Single
Dim PosY As Single
Dim SizeLetra As Integer
Dim No_Carnet As Integer
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
PosX = 0: PosY = 0
SizeLetra = 9
Pagina = 1
No_Carnet = 1
'Iniciamos la impresion
Printer.FontBold = False
sSQL = "SELECT Grupo,Sexo,Cliente,CI_RUC,Direccion " _
       & "FROM Clientes " _
       & "WHERE Grupo = '" & Curso & "' " _
       & "ORDER BY Grupo,Cliente,Sexo DESC "
Select_Adodc AdoAux, sSQL
DataAnchoCampos InicioX, AdoAux, SizeLetra, TipoArialNarrow, Orientacion_Pagina, True
With AdoAux.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
       Printer.Line (PosX + 0.01, PosY + 0.01)-(PosX + 6.5, PosY + 5.5), QBColor(0), B
       PrinterPaint LogoTipo, PosX + 0.1, PosY + 0.2, 2, 1
       Printer.FillColor = QBColor(Rojo)
       Printer.FontBold = True
       Printer.FontSize = 8
       PrinterTexto PosX + 1.8, PosY + 0.1, UCase$(Institucion1)
       Printer.FontSize = 6
       PrinterTexto PosX + 1.8, PosY + 0.5, Institucion2
       Printer.FillColor = QBColor(0)
       Printer.FontSize = 8
       Printer.FontItalic = True
       PrinterTexto PosX + 1.5, PosY + 1.1, "CERTIFICADO DE VOTACION"
       Printer.FontSize = 6
       Printer.FontBold = False
       PrinterTexto PosX + 1.5, PosY + 1.5, UCase$(NombreCiudad) & " - MANABI"
       Printer.FontItalic = False
       Printer.FontBold = True
       Printer.FontSize = 8
       PrinterFields PosX + 0.5, PosY + 2, .Fields("Cliente")
       PrinterFields PosX + 0.5, PosY + 2.8, .Fields("Direccion")
       PrinterFields PosX + 0.5, PosY + 3.6, .Fields("CI_RUC")
       PrinterTexto PosX + 2, PosY + 3.6, Anio_Lectivo
       Printer.FontBold = False
       Printer.FontSize = 7
       PrinterTexto PosX + 0.5, PosY + 2.3, "ESTUDIANTE"
       PrinterTexto PosX + 0.5, PosY + 3.1, "CURSO"
       PrinterTexto PosX + 0.5, PosY + 3.9, "CÓDIGO"
       PrinterTexto PosX + 2, PosY + 3.9, "AÑO LECTIVO"
       
       PrinterTexto PosX + 0.5, PosY + 4.7, Rector
       PrinterTexto PosX + 4, PosY + 4.7, "Estudiante"
       PrinterTexto PosX + 0.5, PosY + 5, TextoRector
       PrinterTexto PosX + 4, PosY + 5, "Firma"
       No_Carnet = No_Carnet + 1
       PosX = PosX + 6.5
       If No_Carnet > 3 Then
          PosX = 0
          PosY = PosY + 5.5
          No_Carnet = 1
       End If
       If PosY >= 25 Then
          Printer.NewPage
          PosY = 0
       End If
      .MoveNext
     Loop
End If
End With
RatonNormal
MensajeEncabData = "    "
Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Actualizar_Notas_Curso(Tabla_Cambiar As String, Campo As String, Codigo_Cliente As String)
    RatonReloj
    If SQL_Server Then
       sSQL = "UPDATE " & Tabla_Cambiar & " " _
            & "SET " & Campo & " = C.Grupo " _
            & "FROM " & Tabla_Cambiar & " As CM, Clientes As C "
    Else
       sSQL = "UPDATE " & Tabla_Cambiar & " As CM, Clientes As C " _
            & "SET CM." & Campo & " = C.Grupo "
    End If
    sSQL = sSQL _
         & "WHERE CM.Item = '" & NumEmpresa & "' " _
         & "AND CM.Periodo = '" & Periodo_Contable & "' " _
         & "AND C.FA <> " & Val(adFalse) & " " _
         & "AND C.Codigo = '" & Codigo_Cliente & "' " _
         & "AND LEN(C.Grupo) = 7 " _
         & "AND CM.Codigo = C.Codigo "
    'MsgBox sSQL
    Ejecutar_SQL_SP sSQL
    RatonNormal
End Sub
