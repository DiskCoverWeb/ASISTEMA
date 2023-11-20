VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form ListClientes 
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11580
   WindowState     =   1  'Minimized
   Begin VB.OptionButton OpcCli 
      Caption         =   "Clientes"
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
      Top             =   105
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.OptionButton OpcCxC 
      Caption         =   "CxC"
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
      Left            =   1575
      TabIndex        =   78
      Top             =   105
      Width           =   1065
   End
   Begin VB.OptionButton OpcTodos 
      Caption         =   "Todos"
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
      Left            =   4830
      TabIndex        =   77
      Top             =   105
      Width           =   1065
   End
   Begin VB.OptionButton OpcCxP 
      Caption         =   "CxP"
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
      TabIndex        =   76
      Top             =   105
      Width           =   1275
   End
   Begin VB.ListBox LstCampos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   4635
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
      Height          =   330
      Left            =   10080
      TabIndex        =   60
      Top             =   6615
      Width           =   750
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5160
      Left            =   105
      TabIndex        =   5
      Top             =   1890
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9102
      _Version        =   393216
      TabOrientation  =   3
      TabHeight       =   882
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1.- Datos Principales"
      TabPicture(0)   =   "ListClie.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label27"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label31"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label25"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label34"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label35"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label26"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblCodigo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label28"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label24"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label36"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label30"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label37"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label38"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label9"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label5"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label11"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtTelefonoS"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtFAX"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtCelular"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtCiudadS"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtCasilla"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtEmail"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtRazonSocial"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "CNacion"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "CProvincia"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtProfesion"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TxtApellidosS"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtActividad"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtDirS"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtNo_Dep"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtDirT"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TxtTelefonoT"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TxtGrupo"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "MBFecha"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TxtCI_RUC"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "MBFechaN"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtNumero"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtLugarTrabS"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "CEstado"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "OpcM"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "OpcF"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TxtLDirs"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "LstProductos"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).ControlCount=   54
      TabCaption(1)   =   "&2.- Datos Secundarios"
      TabPicture(1)   =   "ListClie.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3.- Listado de Beneficiarios"
      TabPicture(2)   =   "ListClie.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DGClientes"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command4 
         Height          =   330
         Left            =   -74895
         Picture         =   "ListClie.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   4725
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   330
         Left            =   -74580
         Picture         =   "ListClie.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4725
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGClientes 
         Bindings        =   "ListClie.frx":0D80
         Height          =   4530
         Left            =   -74895
         TabIndex        =   73
         Top             =   105
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   7990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         WrapCellPointer =   -1  'True
         RowDividerStyle =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "LISTADO DE CLIENTES"
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
      Begin VB.ListBox LstProductos 
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
         Height          =   1620
         Left            =   7245
         TabIndex        =   58
         Top             =   3045
         Width           =   3480
      End
      Begin VB.TextBox TxtLDirs 
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
         Left            =   105
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Text            =   "ListClie.frx":0D9A
         Top             =   3045
         Width           =   7155
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
         TabIndex        =   14
         Top             =   315
         Width           =   1275
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
         TabIndex        =   13
         Top             =   105
         Width           =   1275
      End
      Begin VB.ComboBox CEstado 
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
         Left            =   6405
         TabIndex        =   30
         Text            =   "Soltero"
         Top             =   1365
         Width           =   1275
      End
      Begin VB.TextBox TxtLugarTrabS 
         BeginProperty Font 
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
         MaxLength       =   30
         TabIndex        =   46
         Top             =   1890
         Width           =   2745
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
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   28
         Top             =   1365
         Width           =   1380
      End
      Begin MSMask.MaskEdBox MBFechaN 
         Height          =   330
         Left            =   7665
         TabIndex        =   32
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
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
         Left            =   105
         MaxLength       =   13
         TabIndex        =   7
         Top             =   315
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   9450
         TabIndex        =   54
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   2415
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
      Begin VB.TextBox TxtGrupo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9870
         MaxLength       =   3
         TabIndex        =   36
         Top             =   1365
         Width           =   855
      End
      Begin VB.TextBox TxtTelefonoT 
         BeginProperty Font 
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
         MaxLength       =   20
         TabIndex        =   52
         Top             =   2415
         Width           =   1590
      End
      Begin VB.TextBox TxtDirT 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   50
         Top             =   2415
         Width           =   4425
      End
      Begin VB.TextBox TxtNo_Dep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8925
         MaxLength       =   3
         TabIndex        =   34
         Top             =   1365
         Width           =   960
      End
      Begin VB.TextBox TxtDirS 
         BeginProperty Font 
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
         TabIndex        =   26
         Top             =   1365
         Width           =   4950
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
         Left            =   6405
         MaxLength       =   25
         TabIndex        =   44
         Top             =   1890
         Width           =   1590
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
         Left            =   1785
         MaxLength       =   60
         TabIndex        =   10
         Top             =   315
         Width           =   5265
      End
      Begin VB.TextBox TxtProfesion 
         BeginProperty Font 
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
         MaxLength       =   25
         TabIndex        =   38
         Top             =   1890
         Width           =   1695
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
         Left            =   3465
         TabIndex        =   18
         Text            =   "PICHINCHA"
         Top             =   840
         Width           =   2955
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
         Top             =   840
         Width           =   3375
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
         Left            =   105
         MaxLength       =   50
         TabIndex        =   48
         Top             =   2415
         Width           =   3375
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
         Left            =   3255
         MaxLength       =   50
         TabIndex        =   42
         Top             =   1890
         Width           =   3165
      End
      Begin VB.TextBox TxtCasilla 
         BeginProperty Font 
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
         MaxLength       =   15
         TabIndex        =   40
         Top             =   1890
         Width           =   1485
      End
      Begin VB.TextBox TxtCiudadS 
         BeginProperty Font 
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
         MaxLength       =   20
         TabIndex        =   12
         Top             =   315
         Width           =   2325
      End
      Begin VB.Frame Frame5 
         Caption         =   "Productos Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74895
         TabIndex        =   61
         Top             =   105
         Width           =   10620
         Begin VB.CommandButton Command3 
            Caption         =   "Apertura"
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
            Left            =   1470
            Picture         =   "ListClie.frx":0D9E
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   210
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DGCuentas 
            Bindings        =   "ListClie.frx":11E0
            Height          =   1170
            Left            =   105
            TabIndex        =   66
            Top             =   945
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   2064
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
         Begin VB.CommandButton Command11 
            Caption         =   "Activar"
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
            Left            =   2415
            Picture         =   "ListClie.frx":11F9
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   210
            Width           =   855
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Bloquear"
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
            Left            =   3360
            Picture         =   "ListClie.frx":188F
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   210
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Activar"
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
            Left            =   7875
            Picture         =   "ListClie.frx":1F25
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   210
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Bloquear"
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
            Left            =   8820
            Picture         =   "ListClie.frx":25BB
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   210
            Width           =   855
         End
         Begin MSMask.MaskEdBox MBCuenta 
            Height          =   330
            Left            =   105
            TabIndex        =   63
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   525
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   192
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
            Format          =   "CCCCCCCC-C"
            Mask            =   "########-#"
            PromptChar      =   "0"
         End
         Begin MSMask.MaskEdBox MBTarjeta 
            Height          =   330
            Left            =   5565
            TabIndex        =   68
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   525
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   192
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "CCCCCCC-CCCCCCCC-C"
            Mask            =   "#######-########-#"
            PromptChar      =   "0"
         End
         Begin MSDataGridLib.DataGrid DGTarjetas 
            Bindings        =   "ListClie.frx":2C51
            Height          =   1170
            Left            =   5355
            TabIndex        =   71
            Top             =   945
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   2064
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
         Begin VB.Label Label33 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TARJETA No."
            BeginProperty Font 
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
            TabIndex        =   67
            Top             =   210
            Width           =   2220
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CUENTA No."
            BeginProperty Font 
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
            TabIndex        =   62
            Top             =   210
            Width           =   1275
         End
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
         Left            =   9345
         MaxLength       =   20
         TabIndex        =   24
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox TxtFAX 
         BeginProperty Font 
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
         MaxLength       =   20
         TabIndex        =   22
         Top             =   840
         Width           =   1485
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
         Left            =   6405
         MaxLength       =   20
         TabIndex        =   20
         Top             =   840
         Width           =   1485
      End
      Begin VB.Data DataCuentas 
         Caption         =   "Utilice las Flechas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1260
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   -945
         Width           =   10410
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. DEP."
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
         Left            =   8925
         TabIndex        =   33
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* TELEFONO"
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
         Left            =   6405
         TabIndex        =   19
         Top             =   630
         Width           =   1485
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
         Height          =   225
         Left            =   7980
         TabIndex        =   45
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label Label38 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* NUMERO"
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
         Left            =   5040
         TabIndex        =   27
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Label Label37 
         Caption         =   "* SON ITEM OBLIGATORIOS"
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
         TabIndex        =   59
         Top             =   4725
         Width           =   2955
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRODUCTOS RELACIONADOS"
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
         Left            =   7245
         TabIndex        =   57
         Top             =   2835
         Width           =   3480
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HISTORIAL DE DIRECCIONES"
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
         TabIndex        =   55
         Top             =   2835
         Width           =   7155
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GRUPO #"
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
         Left            =   9870
         TabIndex        =   35
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL"
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
         Left            =   3255
         TabIndex        =   41
         Top             =   1680
         Width           =   3165
      End
      Begin VB.Label Label22 
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
         Left            =   7665
         TabIndex        =   31
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* C.I./R.U.C."
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
         TabIndex        =   6
         Top             =   105
         Width           =   1695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA APE."
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
         TabIndex        =   53
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label Label10 
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
         Height          =   225
         Left            =   7875
         TabIndex        =   51
         Top             =   2205
         Width           =   1590
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DIRECCION DEL TRABAJO"
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
         TabIndex        =   49
         Top             =   2205
         Width           =   4425
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* CIUDAD"
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
         Left            =   7035
         TabIndex        =   11
         Top             =   105
         Width           =   2325
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* DIRECCION"
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
         TabIndex        =   25
         Top             =   1155
         Width           =   4950
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ACTIVIDAD"
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
         Left            =   6405
         TabIndex        =   43
         Top             =   1680
         Width           =   1590
      End
      Begin VB.Label LblCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   225
         Left            =   5670
         TabIndex        =   9
         Top             =   105
         Width           =   1380
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* APELLIDOS Y NOMBRES"
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
         Left            =   1785
         TabIndex        =   8
         Top             =   105
         Width           =   3900
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
         Height          =   225
         Left            =   105
         TabIndex        =   37
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* PROVINCIA"
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
         Left            =   3465
         TabIndex        =   17
         Top             =   630
         Width           =   2955
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* NACIONALIDAD"
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
         TabIndex        =   15
         Top             =   630
         Width           =   3375
      End
      Begin VB.Label Label25 
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
         Height          =   225
         Left            =   105
         TabIndex        =   47
         Top             =   2205
         Width           =   3375
      End
      Begin VB.Label Label31 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EST.CIVIL"
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
         Left            =   6405
         TabIndex        =   29
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CASILLA POS."
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
         TabIndex        =   39
         Top             =   1680
         Width           =   1485
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
         Height          =   225
         Left            =   9345
         TabIndex        =   23
         Top             =   630
         Width           =   1380
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* FAX"
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
         Left            =   7875
         TabIndex        =   21
         Top             =   630
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtCIRUC 
      BeginProperty Font 
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
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1470
      Width           =   4635
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "ListClie.frx":2C6B
      DataSource      =   "AdoListCtas"
      Height          =   1155
      Left            =   4830
      TabIndex        =   4
      Top             =   630
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   2037
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
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   420
      Top             =   3780
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
      Left            =   420
      Top             =   3465
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
      Left            =   420
      Top             =   3150
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
      Left            =   420
      Top             =   4095
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
      Left            =   420
      Top             =   4410
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
      Top             =   2835
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
   Begin VB.Label Label18 
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
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   4635
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL CLIENTE"
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
      Left            =   4830
      TabIndex        =   3
      Top             =   420
      Width           =   6630
   End
End
Attribute VB_Name = "ListClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub DatosNuevos()
   LblCodigo.Caption = "Ninguno"
   'LblEstado.Caption = ""
   TxtApellidosS.Text = ""
   TxtCI_RUC.Text = ""
   TxtTelefonoS.Text = ""
   TxtDirS.Text = ""
   TxtNo_Dep.Text = ""
   TxtGrupo.Text = ""
   TxtCiudadS.Text = ""
   TxtLugarTrabS.Text = ""
   TxtDirT.Text = ""
   TxtTelefonoT.Text = ""
   TxtProfesion.Text = ""
   TxtRazonSocial.Text = ""
   TxtFAX.Text = ""
   TxtCasilla.Text = ""
   TxtActividad.Text = ""
   TxtEmail.Text = ""
   TxtApellidosS.Enabled = True
   TxtRazonSocial.Enabled = True
End Sub

Public Sub ListarClientes(Optional LlenarCliente As Boolean, _
                          Optional OrdenarColumna As Integer, _
                          Optional Descendente As Boolean)
  If OpcCli.value Then
     sSQL = "SELECT Cliente,T,Fecha_N,TD,CI_RUC,Direccion,DirNumero,Grupo,Telefono,Celular," _
          & "FAX,Prov,Ciudad,Est_Civil,Sexo,Profesion,Actividad,Email,C.Codigo,Fecha," _
          & "Representante,Casilla,Lugar_Trabajo,DireccionT,TelefonoT,No_Dep " _
          & "FROM Clientes As C " _
          & "WHERE Cliente <> '.' " _
          & "AND FA <> " & Val(adFalse) & " "
  ElseIf OpcCxC.value Then
     sSQL = "SELECT Cliente,T,Fecha_N,TD,CI_RUC,Direccion,DirNumero,Grupo,Telefono,Celular," _
          & "FAX,Prov,Ciudad,Est_Civil,Sexo,Profesion,Actividad,Email,C.Codigo,Fecha," _
          & "Representante,Casilla,Lugar_Trabajo,DireccionT,TelefonoT,No_Dep " _
          & "FROM Clientes As C,Catalogo_CxCxP As CCP " _
          & "WHERE Cliente <> '.' " _
          & "AND CCP.Item = '" & NumEmpresa & "' " _
          & "AND CCP.Periodo = '" & Periodo_Contable & "' " _
          & "AND CCP.TC = 'C' " _
          & "AND CCP.Codigo = C.Codigo "
  ElseIf OpcCxP.value Then
     sSQL = "SELECT Cliente,T,Fecha_N,TD,CI_RUC,Direccion,DirNumero,Grupo,Telefono,Celular," _
          & "FAX,Prov,Ciudad,Est_Civil,Sexo,Profesion,Actividad,Email,C.Codigo,Fecha," _
          & "Representante,Casilla,Lugar_Trabajo,DireccionT,TelefonoT,No_Dep " _
          & "FROM Clientes As C,Catalogo_CxCxP As CCP " _
          & "WHERE Cliente <> '.' " _
          & "AND CCP.Item = '" & NumEmpresa & "' " _
          & "AND CCP.Periodo = '" & Periodo_Contable & "' " _
          & "AND CCP.TC = 'P' " _
          & "AND CCP.Codigo = C.Codigo "
  Else
     sSQL = "SELECT Cliente,T,Fecha_N,TD,CI_RUC,Direccion,DirNumero,Grupo,Telefono,Celular," _
          & "FAX,Prov,Ciudad,Est_Civil,Sexo,Profesion,Actividad,Email,C.Codigo,Fecha," _
          & "Representante,Casilla,Lugar_Trabajo,DireccionT,TelefonoT,No_Dep " _
          & "FROM Clientes As C " _
          & "WHERE Cliente <> '.' "
  End If
  sSQL = sSQL & "GROUP BY Cliente,T,Fecha_N,TD,CI_RUC,Direccion,DirNumero,Grupo,Telefono," _
       & "Celular,FAX,Prov,Ciudad,Est_Civil,Sexo,Profesion,Actividad,Email,C.Codigo," _
       & "Fecha,Representante,Casilla,Lugar_Trabajo,DireccionT,TelefonoT,No_Dep "
  If OrdenarColumna > 0 Then
     sSQL = sSQL & "ORDER BY " & DGClientes.Columns(OrdenarColumna).DataField & " "
     If Descendente Then sSQL = sSQL & "DESC "
  Else
     sSQL = sSQL & "ORDER BY Cliente "
     If Descendente Then sSQL = sSQL & "DESC "
  End If
  SelectDBCombo DCCliente, AdoListCtas, sSQL, "Cliente"
  If LlenarCliente Then
     LstCampos.Clear
     With AdoListCtas.Recordset
      For I = 0 To .Fields.Count - 1
          LstCampos.AddItem .Fields(I).Name
      Next I
     End With
     LstCampos.Text = "Cliente"
  End If
  DGClientes.AllowUpdate = False
  Label2.Caption = " NOMBRE DEL CLIENTE" & Space(30) & "Total Clientes: " & Format(AdoListCtas.Recordset.RecordCount, "000000")
  DCCliente.SetFocus
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  DatosNuevos
  LstProductos.Clear
  LblCodigo.Caption = "Ninguno"
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & TextoBusqueda & "'")
       If Not .EOF Then
          If .Fields("T") = "N" Then LstProductos.AddItem "Activo" Else LstProductos.AddItem "Inactivo"
          MBFecha.Text = .Fields("Fecha")
          MBFechaN.Text = .Fields("Fecha_N")
          LblCodigo.Caption = .Fields("Codigo")
          TxtApellidosS.Text = .Fields("Cliente")
          DCCliente.Text = .Fields("Cliente")
          TxtCI_RUC.Text = .Fields("CI_RUC")
          TxtProfesion.Text = .Fields("Profesion")
          TxtActividad.Text = .Fields("Actividad")
          TxtRazonSocial.Text = .Fields("Representante")
          TxtCasilla.Text = .Fields("Casilla")
          TxtEmail.Text = .Fields("Email")
          TxtTelefonoS.Text = .Fields("Telefono")
          TxtFAX.Text = .Fields("FAX")
          TxtCelular.Text = .Fields("Celular")
          TxtDirS.Text = .Fields("Direccion")
          TxtNumero.Text = .Fields("DirNumero")
          TxtCiudadS.Text = .Fields("Ciudad")
          TxtLugarTrabS.Text = .Fields("Lugar_Trabajo")
          TxtDirT.Text = .Fields("DireccionT")
          TxtTelefonoT.Text = .Fields("TelefonoT")
          TxtNo_Dep.Text = .Fields("No_Dep")
          TxtGrupo.Text = .Fields("Grupo")
          For I = 0 To CEstado.ListCount - 1
           If .Fields("Est_Civil") = Mid(CEstado.List(I), 1, 1) Then
               CEstado.Text = CEstado.List(I)
           End If
          Next I
          If .Fields("Sexo") = "M" Then OpcM.value = True Else OpcF.value = True
          Label6.Caption = "* C.I./R.U.C.   [" & .Fields("TD") & "]"
          For I = 0 To CProvincia.ListCount - 1
           If .Fields("Prov") = Mid(CProvincia.List(I), 1, 2) Then
               CProvincia.Text = CProvincia.List(I)
           End If
          Next I
'''          For I = 0 To CNacion.ListCount - 1
'''           If .Fields("Pais") = Mid(CNacion.List(I), 1, 3) Then
'''               CNacion.Text = CNacion.List(I)
'''           End If
'''          Next I
          TxtApellidosS.Enabled = False
          TxtRazonSocial.Enabled = False
          sSQL = "SELECT * " _
               & "FROM Clientes_Libretas " _
               & "WHERE Codigo = '" & LblCodigo.Caption & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "ORDER BY Fecha "
          SelectDataGrid DGCuentas, AdoCuentas, sSQL
          If AdoCuentas.Recordset.RecordCount > 0 Then
             Do While Not AdoCuentas.Recordset.EOF
                MBCuenta.Text = AdoCuentas.Recordset.Fields("Cuenta_No")
                LstProductos.AddItem "Cta. Ahorro No. " & MBCuenta.Text
                AdoCuentas.Recordset.MoveNext
             Loop
          End If
          sSQL = "SELECT * " _
               & "FROM Catalogo_CxCxP " _
               & "WHERE Codigo = '" & LblCodigo.Caption & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "ORDER BY TC,Cta "
          SelectData AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             Do While Not AdoAux.Recordset.EOF
                LstProductos.AddItem "Cta. Contable (" & AdoAux.Recordset.Fields("TC") & "): " & AdoAux.Recordset.Fields("Cta")
                AdoAux.Recordset.MoveNext
             Loop
          End If
          sSQL = "SELECT * " _
               & "FROM Catalogo_Rol_Pagos " _
               & "WHERE Codigo = '" & LblCodigo.Caption & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          SelectData AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then LstProductos.AddItem "Asignado a Rol de Pago"
          sSQL = "SELECT * " _
               & "FROM Prestamos " _
               & "WHERE Cuenta_No = '" & LblCodigo.Caption & "' " _
               & "AND TP = 'SUSC' " _
               & "ORDER BY Fecha,Credito_No "
          SelectData AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             Do While Not AdoAux.Recordset.EOF
                LstProductos.AddItem "Suscripción: [" & AdoAux.Recordset.Fields("Fecha") & "] " & AdoAux.Recordset.Fields("Credito_No")
                AdoAux.Recordset.MoveNext
             Loop
          End If
          sSQL = "SELECT * " _
               & "FROM Clientes_Datos_Extras " _
               & "WHERE Codigo = '" & LblCodigo.Caption & "' " _
               & "ORDER BY Fecha "
          SelectAdodc AdoAux, sSQL
          TxtLDirs.Text = ""
          With AdoAux.Recordset
           If .RecordCount > 0 Then
               Do While Not .EOF
                  TxtLDirs.Text = TxtLDirs.Text _
                                & .Fields("Fecha") & ": " & .Fields("Ciudad") & ", " _
                                & .Fields("Direccion") & ". " & .Fields("Telefono") & vbCrLf
                 .MoveNext
               Loop
           End If
          End With
          Mifecha = PrimerDiaMes(FechaSistema)
          Dia = Day(Mifecha)
          Mes = Month(Mifecha)
          Anio = Year(Mifecha)
          FechaIni = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anio, "0000")
          FechaFin = FechaSistema
          Total = 0: Saldo = 0: Contador = 1
       Else
          MsgBox "No Existe"
          DatosNuevos
       End If
   Else
     MsgBox "No Existe"
     DatosNuevos
   End If
  End With
End Sub

Private Sub CEstado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Command1_Click()
  Unload ListClientes
End Sub

Private Sub Command10_Click()
  Mensajes = "Esta seguro de desea bloquear la Cuenta No. " & MBCuenta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Clientes_Libretas " _
          & "SET T = 'A' " _
          & "WHERE Cuenta_No = '" & MBCuenta.Text & "' "
     ConectarAdoExecute sSQL
  End If
  ListarCuenta DCCliente.Text
End Sub

Private Sub Command11_Click()
  Mensajes = "Esta seguro de desea desbloquear la Cuenta No. " & MBCuenta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Clientes_Libretas " _
          & "SET T = 'N' " _
          & "WHERE Cuenta_No = '" & MBCuenta.Text & "' "
     ConectarAdoExecute sSQL
  End If
  ListarCuenta DCCliente.Text
End Sub

Private Sub Command2_Click()
  ListarClientes False, Pagina, True
End Sub

Private Sub Command3_Click()
  Imprimir_Apertura MBCuenta.Text
End Sub

Private Sub Command4_Click()
  ListarClientes False, Pagina
End Sub

Private Sub Command8_Click()
  Mensajes = "Esta seguro de desea desbloquear la Tarjeta No. " & MBTarjeta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Tarjetas " _
          & "SET T = 'N' " _
          & "WHERE Tarjeta_No = '" & MBTarjeta.Text & "' "
     'ConectarAdoExecute sSQL
  End If
  ListarCuenta DCCliente.Text
End Sub

Private Sub Command9_Click()
  Mensajes = "Esta seguro de desea bloquear la Tarjeta No. " & MBTarjeta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Tarjetas " _
          & "SET T = 'B' " _
          & "WHERE Tarjeta_No = '" & MBTarjeta.Text & "' "
     'ConectarAdoExecute sSQL
  End If
  ListarCuenta DCCliente.Text
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente.Text
  TipoDoc = "M"
End Sub

Private Sub DGClientes_HeadClick(ByVal ColIndex As Integer)
  Pagina = ColIndex
End Sub

Private Sub DGClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  Select Case KeyCode
    Case vbKeyF1
         DGClientes.Visible = False
         GenerarDataTexto ListClientes, AdoListCtas
         DGClientes.Visible = True
    Case vbKeyF11
         DGClientes.Visible = False
         With AdoListCtas.Recordset
          If .RecordCount > 0 Then
              Contador = 0
             .MoveFirst
              Do While Not .EOF
                 Contador = Contador + 1
                 ListClientes.Caption = Contador & " / " & .RecordCount
                 Cadena = CompilarRUC_CI(.Fields("CI_RUC"))
                 Cadena1 = Cadena
                 'MsgBox Cadena
                 If Cadena = Ninguno Then Cadena = "90X" & Format(Contador, "000000")
                 Si_No = False
                 If .Fields("CI_RUC") <> Trim(Cadena) Then
                    .Fields("CI_RUC") = Trim(Cadena)
                     Si_No = True
                 End If
                 NombreCliente = Trim(UCase(.Fields("Cliente")))
                 DireccionCli = Trim(UCase(.Fields("Direccion")))
                 If DireccionCli = Ninguno Then DireccionCli = "S/D"
                 If Mid(NombreCliente, 1, 1) = "." Then NombreCliente = Mid(NombreCliente, 2, Len(NombreCliente))
                 If Mid(NombreCliente, Len(NombreCliente), 1) = "." Then NombreCliente = Mid(NombreCliente, 1, Len(NombreCliente) - 1)
                 
                 If Mid(DireccionCli, 1, 1) = "." Then DireccionCli = Mid(DireccionCli, 2, Len(DireccionCli))
                 If Mid(DireccionCli, Len(DireccionCli), 1) = "." Then DireccionCli = Mid(DireccionCli, 1, Len(DireccionCli) - 1)
                
                .Fields("Cliente") = Trim(NombreCliente)
                .Fields("Direccion") = Trim(DireccionCli)
                .Fields("T") = Normal
                .Fields("Ciudad") = UCase(.Fields("Ciudad"))
                 Select Case Mid(Cadena, 3, 1)
                   Case "0" To "6", "8": If Len(Cadena) = 13 Then Cadena = "R" Else Cadena = "C"
                   Case "7": Cadena = "P"
                   Case "9": Cadena = "R"
                   Case Else: Cadena = "O"
                 End Select
                 If Len(Cadena1) < 10 Then Cadena = "O"
                 If Val(Mid(Cadena1, 1, 1)) >= 3 Then Cadena = "O"
                 If .Fields("TD") <> Cadena Then
                    .Fields("TD") = Cadena
                     Si_No = True
                 End If
                 If Len(.Fields("Telefono")) <= 2 Then .Fields("Telefono") = "000000000"
                 If Len(.Fields("Celular")) <= 2 Then .Fields("Celular") = "000000000"
                 If Len(.Fields("FAX")) <= 2 Then .Fields("FAX") = "000000000"
                 If Len(.Fields("Direccion")) <= 2 Then .Fields("Direccion") = "S/D"
                 If Len(.Fields("DirNumero")) <= 2 Then .Fields("DirNumero") = "S/N"
                .Update
                .MoveNext
              Loop
          End If
         End With
         DGClientes.Visible = True
    Case vbKeyF12
  End Select
  If CtrlDown And KeyCode = vbKeyF5 Then DGClientes.AllowUpdate = True
  If CtrlDown And KeyCode = vbKeyP Then
     DGClientes.Visible = False
     Imprimir_Clientes AdoListCtas
     'ImprimirAdodc AdoListCtas, True, 2, 8
     DGClientes.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyB Then
     'CampoBusqueda
     BuscarDatos DGClientes, AdoListCtas
  End If
End Sub

Private Sub DGCuentas_DblClick()
   If AdoCuentas.Recordset.RecordCount > 0 Then
      MBCuenta.Text = AdoCuentas.Recordset.Fields("Cuenta_No")
   Else
      MsgBox "No existe cuenta"
   End If
End Sub

Private Sub DGTarjetas_DblClick()
   If AdoTarjetas.Recordset.RecordCount > 0 Then
      MBtarjetas.Text = AdoTarjetas.Recordset.Fields("Tarjeta_No")
   Else
      MsgBox "No existe Tarjeta"
   End If
End Sub

Private Sub Form_Activate()
  ListClientes.Caption = "LISTADO DE CLIENTES"
  LblCodigo.Caption = "Ninguno"
  CEstado.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Estado_Civil " _
       & "WHERE Estado<>'000' "
  SelectData AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CEstado.Text = AdoAux.Recordset.Fields("Estado")
     Do While Not AdoAux.Recordset.EOF
        CEstado.AddItem AdoAux.Recordset.Fields("Estado")
        AdoAux.Recordset.MoveNext
     Loop
  End If
  CNacion.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Nacionalidad " _
       & "WHERE Pais<>'000' " _
       & "ORDER BY Pais "
  SelectData AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Codigo = AdoAux.Recordset.Fields("Pais") & " " _
            & AdoAux.Recordset.Fields("Paises")
     CNacion.Text = Codigo
     Do While Not AdoAux.Recordset.EOF
        Codigo = AdoAux.Recordset.Fields("Pais") & " " _
               & AdoAux.Recordset.Fields("Paises")
        CNacion.AddItem Codigo
        AdoAux.Recordset.MoveNext
     Loop
  End If
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Provincia " _
       & "WHERE Prov<>'00' " _
       & "ORDER BY Prov "
  SelectData AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Codigo = AdoAux.Recordset.Fields("Prov") & "  " _
            & AdoAux.Recordset.Fields("Provincia")
     CProvincia.Text = Codigo
     Do While Not AdoAux.Recordset.EOF
        Codigo = AdoAux.Recordset.Fields("Prov") & "  " _
               & AdoAux.Recordset.Fields("Provincia")
        CProvincia.AddItem Codigo
        AdoAux.Recordset.MoveNext
     Loop
  End If
  ListarClientes True
  RatonNormal
  ListClientes.WindowState = vbMaximized
  If Nuevo Then
     TxtApellidosS.Text = NombreCliente
     LblCodigo.Caption = "Ninguno"
     TxtGrupo.Text = NumEmpresa
     TxtApellidosS.SetFocus
  Else
     ListarCuenta DCCliente.Text
     DCCliente.SetFocus
  End If
End Sub

Private Sub Form_Deactivate()
  ListClientes.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   CentrarForm ListClientes
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
End Sub

Private Sub OpcCli_Click()
  ListarClientes False, Pagina, True
End Sub

Private Sub OpcCxC_Click()
  ListarClientes False, Pagina, True
End Sub

Private Sub OpcCxP_Click()
  ListarClientes False, Pagina, True
End Sub

Private Sub OpcTodos_Click()
  ListarClientes False, Pagina, True
End Sub
