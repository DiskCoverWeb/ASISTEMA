VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Apertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11730
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
      Height          =   960
      Left            =   10605
      Picture         =   "Apertura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5355
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
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
      Height          =   960
      Left            =   10605
      Picture         =   "Apertura.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   4410
      Width           =   1065
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
      Height          =   960
      Left            =   10605
      Picture         =   "Apertura.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3465
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Si&g."
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
      Left            =   11130
      Picture         =   "Apertura.frx":0D2E
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   2625
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ant."
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
      Left            =   10605
      Picture         =   "Apertura.frx":1170
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2625
      Width           =   540
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mov. de:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   10605
      TabIndex        =   73
      Top             =   0
      Width           =   1065
      Begin VB.OptionButton Option3 
         Caption         =   "&90 días"
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
         Left            =   105
         Picture         =   "Apertura.frx":15B2
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&60 días"
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
         Left            =   105
         Picture         =   "Apertura.frx":19F4
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   945
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&30 días"
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
         Left            =   105
         Picture         =   "Apertura.frx":1E36
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   105
      TabIndex        =   70
      Top             =   840
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   10583
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DATOS DE LA CUENTA"
      TabPicture(0)   =   "Apertura.frx":2278
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label27"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label26"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label28"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "LabelEstado"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label22"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label23"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label30"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtDirS"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtNombresS"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CheckConyugue"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtRazonSocial"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtApellidosS"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "MBoxRUCS"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtDirT"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtTelefonoS"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtFAX"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtTelefonoT"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtLugarTrabS"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtCasilla"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Command6"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "MBoxCI"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtProfesion"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtCiudadS"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TxtEstCiv"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtNo_Dep"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtNoSoc"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtActividad"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtSector"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TxtArea"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "DGListCtas"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "FrameConyugue"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "MOVIMIENTO/INTERESES"
      TabPicture(1)   =   "Apertura.frx":2294
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGInt"
      Tab(1).Control(1)=   "DGMovCta"
      Tab(1).Control(2)=   "AdoIntereses"
      Tab(1).Control(3)=   "AdoMovCta"
      Tab(1).Control(4)=   "LabelTotInt"
      Tab(1).Control(5)=   "Label24"
      Tab(1).Control(6)=   "LabelPromedio"
      Tab(1).Control(7)=   "Label29"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "CREDITOS OTORGADOS"
      TabPicture(2)   =   "Apertura.frx":22B0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "AdoCreditos"
      Tab(2).Control(1)=   "DGCreditos"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "TARJETAS DE DEBITOS"
      TabPicture(3)   =   "Apertura.frx":22CC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DGTarjeta"
      Tab(3).Control(1)=   "Command9"
      Tab(3).Control(2)=   "Command8"
      Tab(3).Control(3)=   "Command7"
      Tab(3).Control(4)=   "MBoxTarjeta"
      Tab(3).Control(5)=   "MBoxFechaT"
      Tab(3).Control(6)=   "Label32"
      Tab(3).Control(7)=   "Label31"
      Tab(3).ControlCount=   8
      Begin VB.Frame FrameConyugue 
         Caption         =   " DATOS DEL CONYUGUE "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   5145
         TabIndex        =   50
         Top             =   3360
         Visible         =   0   'False
         Width           =   5160
         Begin VB.TextBox TxtLugarTrabC 
            BeginProperty Font 
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
            MaxLength       =   30
            TabIndex        =   61
            Top             =   1470
            Width           =   3270
         End
         Begin VB.TextBox TxtTelefonoC 
            BeginProperty Font 
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
            TabIndex        =   60
            Top             =   1155
            Width           =   3270
         End
         Begin MSMask.MaskEdBox MBoxCIC 
            Height          =   330
            Left            =   1785
            TabIndex        =   59
            ToolTipText     =   "Formato de Fecha: DD/MM/AA"
            Top             =   840
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "CCCCCCCCC-C"
            Mask            =   "#########-#"
            PromptChar      =   "0"
         End
         Begin VB.TextBox TxtApellidosC 
            BeginProperty Font 
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
            MaxLength       =   25
            TabIndex        =   58
            Top             =   525
            Width           =   3270
         End
         Begin VB.TextBox TxtNombresC 
            BeginProperty Font 
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
            MaxLength       =   25
            TabIndex        =   57
            Top             =   210
            Width           =   3270
         End
         Begin VB.TextBox TxtDirC 
            BeginProperty Font 
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
            TabIndex        =   62
            Top             =   2100
            Width           =   4950
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Direccion del Trabajo"
            BeginProperty Font 
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
            Top             =   1785
            Width           =   4950
         End
         Begin VB.Label Label16 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Lugar de Trabajo"
            BeginProperty Font 
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
            TabIndex        =   55
            Top             =   1470
            Width           =   1695
         End
         Begin VB.Label Label15 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Telefono"
            BeginProperty Font 
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
            TabIndex        =   54
            Top             =   1155
            Width           =   1695
         End
         Begin VB.Label Label17 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Identificacion"
            BeginProperty Font 
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
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Apellidos"
            BeginProperty Font 
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
            TabIndex        =   52
            Top             =   525
            Width           =   1695
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Nombres"
            BeginProperty Font 
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
            TabIndex        =   51
            Top             =   210
            Width           =   1695
         End
      End
      Begin MSDataGridLib.DataGrid DGCreditos 
         Bindings        =   "Apertura.frx":22E8
         Height          =   5160
         Left            =   -74895
         TabIndex        =   92
         Top             =   420
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   9102
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSAdodcLib.Adodc AdoCreditos 
         Height          =   330
         Left            =   -74895
         Top             =   5565
         Width           =   3585
         _ExtentX        =   6324
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
         Caption         =   "Adodc1"
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
      Begin MSDataGridLib.DataGrid DGTarjeta 
         Bindings        =   "Apertura.frx":2302
         Height          =   4635
         Left            =   -74895
         TabIndex        =   91
         Top             =   1260
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8176
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
      Begin MSDataGridLib.DataGrid DGInt 
         Bindings        =   "Apertura.frx":231C
         Height          =   5160
         Left            =   -67650
         TabIndex        =   90
         Top             =   420
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   9102
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGMovCta 
         Bindings        =   "Apertura.frx":2337
         Height          =   5160
         Left            =   -69330
         TabIndex        =   87
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   9102
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid DGListCtas 
         Bindings        =   "Apertura.frx":234F
         Height          =   1485
         Left            =   105
         TabIndex        =   86
         Top             =   3990
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   2619
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.CommandButton Command9 
         Caption         =   "&Bloquear Tarjeta"
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
         Left            =   -68175
         Picture         =   "Apertura.frx":2369
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   420
         Width           =   1380
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Desbloquear Tarjeta"
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
         Left            =   -69645
         Picture         =   "Apertura.frx":29FF
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   420
         Width           =   1380
      End
      Begin VB.CommandButton Command7 
         Caption         =   "G&rabar Tarjeta"
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
         Left            =   -71115
         Picture         =   "Apertura.frx":3095
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   420
         Width           =   1380
      End
      Begin VB.TextBox TxtArea 
         BeginProperty Font 
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
         MaxLength       =   5
         TabIndex        =   34
         Top             =   3570
         Width           =   645
      End
      Begin VB.TextBox TxtSector 
         BeginProperty Font 
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
         MaxLength       =   25
         TabIndex        =   33
         Top             =   3255
         Width           =   3690
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   32
         Top             =   2955
         Width           =   3690
      End
      Begin VB.TextBox TxtNoSoc 
         BeginProperty Font 
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
         MaxLength       =   3
         TabIndex        =   31
         Top             =   2640
         Width           =   1170
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   30
         Top             =   2640
         Width           =   1380
      End
      Begin VB.TextBox TxtEstCiv 
         BeginProperty Font 
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
         MaxLength       =   15
         TabIndex        =   29
         Top             =   2325
         Width           =   3690
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
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2010
         Width           =   3690
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
         Left            =   1470
         MaxLength       =   25
         TabIndex        =   27
         Top             =   1695
         Width           =   3690
      End
      Begin MSMask.MaskEdBox MBoxCI 
         Height          =   330
         Left            =   3675
         TabIndex        =   26
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1365
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "CCCCCCCCC-C"
         Mask            =   "#########-#"
         PromptChar      =   "0"
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Desbloquear Cierre de Cuenta"
         BeginProperty Font 
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
         TabIndex        =   77
         Top             =   5565
         Width           =   2745
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
         Left            =   7140
         MaxLength       =   15
         TabIndex        =   48
         Top             =   2955
         Width           =   3165
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
         Left            =   7140
         MaxLength       =   30
         TabIndex        =   47
         Top             =   2640
         Width           =   3165
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
         Left            =   7140
         MaxLength       =   20
         TabIndex        =   46
         Top             =   2325
         Width           =   3165
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
         Left            =   7140
         MaxLength       =   20
         TabIndex        =   45
         Top             =   2010
         Width           =   3165
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
         Left            =   7140
         MaxLength       =   20
         TabIndex        =   44
         Top             =   1695
         Width           =   3165
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
         Left            =   5145
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1380
         Width           =   5160
      End
      Begin MSMask.MaskEdBox MBoxRUCS 
         Height          =   330
         Left            =   1470
         TabIndex        =   25
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1380
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
         Mask            =   "#########-#-###"
         PromptChar      =   "0"
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
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   24
         Top             =   1065
         Width           =   3690
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
         Left            =   1470
         MaxLength       =   49
         TabIndex        =   23
         Top             =   750
         Width           =   3690
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
      Begin VB.CheckBox CheckConyugue 
         Caption         =   "Ingresar Datos del Conyugue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2205
         TabIndex        =   49
         Top             =   3675
         Width           =   2850
      End
      Begin VB.TextBox TxtNombresS 
         BeginProperty Font 
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
         MaxLength       =   30
         TabIndex        =   22
         Top             =   435
         Width           =   3690
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
         Left            =   5145
         MaxLength       =   50
         TabIndex        =   36
         Top             =   735
         Width           =   5160
      End
      Begin MSMask.MaskEdBox MBoxTarjeta 
         Height          =   435
         Left            =   -73425
         TabIndex        =   82
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   735
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   767
         _Version        =   393216
         ForeColor       =   192
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
      Begin MSMask.MaskEdBox MBoxFechaT 
         Height          =   435
         Left            =   -74895
         TabIndex        =   83
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   735
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
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
      Begin MSAdodcLib.Adodc AdoIntereses 
         Height          =   330
         Left            =   -67440
         Top             =   5250
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Intereses"
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
      Begin MSAdodcLib.Adodc AdoMovCta 
         Height          =   330
         Left            =   -74895
         Top             =   5565
         Width           =   4110
         _ExtentX        =   7250
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
         Caption         =   "MovCta"
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
      Begin VB.Label LabelTotInt 
         Alignment       =   1  'Right Justify
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
         Left            =   -66810
         TabIndex        =   88
         Top             =   5565
         Width           =   2115
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -67650
         TabIndex        =   89
         Top             =   5565
         Width           =   855
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA APERT"
         BeginProperty Font 
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
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label31 
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
         Left            =   -73425
         TabIndex        =   84
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Area Geog."
         BeginProperty Font 
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
         TabIndex        =   81
         Top             =   3570
         Width           =   1380
      End
      Begin VB.Label Label23 
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
         Left            =   105
         TabIndex        =   80
         Top             =   3255
         Width           =   1380
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. Socios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2835
         TabIndex        =   21
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " C.I."
         BeginProperty Font 
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
         TabIndex        =   15
         Top             =   1365
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label LabelPromedio 
         Alignment       =   1  'Right Justify
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
         Left            =   -69540
         TabIndex        =   78
         Top             =   5565
         Width           =   1905
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROMEDIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70695
         TabIndex        =   79
         Top             =   5565
         Width           =   1170
      End
      Begin VB.Label LabelEstado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombres"
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
         Left            =   105
         TabIndex        =   76
         Top             =   5565
         Width           =   2115
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Direccion del Socio"
         BeginProperty Font 
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
         TabIndex        =   35
         Top             =   420
         Width           =   5160
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Direccion del Trabajo"
         BeginProperty Font 
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
         TabIndex        =   37
         Top             =   1050
         Width           =   5160
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Actividad"
         BeginProperty Font 
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
         TabIndex        =   20
         Top             =   2955
         Width           =   1380
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. Dep."
         BeginProperty Font 
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
         Top             =   2640
         Width           =   1380
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Estado Civil"
         BeginProperty Font 
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
         Top             =   2325
         Width           =   1380
      End
      Begin VB.Label Label8 
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
         Left            =   105
         TabIndex        =   17
         Top             =   2010
         Width           =   1380
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Profesion"
         BeginProperty Font 
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
         Top             =   1695
         Width           =   1380
      End
      Begin VB.Label Label6 
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
         Left            =   105
         TabIndex        =   14
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Apellidos"
         BeginProperty Font 
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
         TabIndex        =   13
         Top             =   1065
         Width           =   1380
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Representante"
         BeginProperty Font 
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
         TabIndex        =   12
         Top             =   750
         Width           =   1380
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Casilla_Postal"
         BeginProperty Font 
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
         TabIndex        =   43
         Top             =   2955
         Width           =   2010
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Lugar de Trabajo"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   2640
         Width           =   2010
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telefono del Trabajo"
         BeginProperty Font 
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
         TabIndex        =   41
         Top             =   2325
         Width           =   2010
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FAX del Socio"
         BeginProperty Font 
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
         TabIndex        =   40
         Top             =   2010
         Width           =   2010
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Telefono del Socio"
         BeginProperty Font 
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
         TabIndex        =   39
         Top             =   1695
         Width           =   2010
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombres"
         BeginProperty Font 
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
         TabIndex        =   11
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Apertura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3045
      TabIndex        =   3
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton OpcCaja 
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton OpcCred 
         Caption         =   "Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   5
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   2850
      Begin VB.OptionButton OpcJ 
         Caption         =   "Juridica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   2
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton OpcN 
         Caption         =   "Natural"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   8820
      TabIndex        =   10
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   315
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      _Version        =   393216
      ForeColor       =   192
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
   Begin VB.CheckBox CheckME 
      Caption         =   "Moneda Extranjera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5670
      TabIndex        =   6
      Top             =   105
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   435
      Left            =   7140
      TabIndex        =   9
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   315
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1680
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   210
      Top             =   1365
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
      Left            =   210
      Top             =   1050
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema\EMPRESA\SQLs\DiskCove.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Sistema\EMPRESA\SQLs\DiskCove.mdb;Persist Security Info=False"
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
      Left            =   210
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
   Begin VB.Label Label2 
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
      Height          =   225
      Left            =   8820
      TabIndex        =   8
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA APERT."
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
      Left            =   7140
      TabIndex        =   7
      Top             =   105
      Width           =   1590
   End
End
Attribute VB_Name = "Apertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PersonalJuridica(Personal As Boolean)
  If Personal Then
     TxtRazonSocial.Text = Ninguno
     TxtActividad.Text = Ninguno
     Label20.Visible = False
     Label26.Visible = True
     Label28.Visible = False
     Label11.Visible = True
     Label3.Caption = " Nombres"
     TxtActividad.Visible = False
     TxtNo_Dep.Visible = True
     TxtProfesion.Visible = True
     MBoxCI.Visible = False
  Else
     TxtProfesion.Text = Ninguno
     Label20.Visible = True
     Label28.Visible = True
     Label26.Visible = False
     Label11.Visible = False
     Label3.Caption = " Razon Social"
     TxtActividad.Visible = True
     TxtNo_Dep.Visible = False
     TxtProfesion.Visible = False
     MBoxCI.Visible = True
  End If
End Sub

Public Sub ListarCuenta(Cuenta_No As String)
   LabelEstado.Caption = ""
   CheckME.Value = 0
   TxtNombresS.Text = ""
   TxtApellidosS.Text = ""
   MBoxRUCS.Text = "000000000-0-000"
   MBoxCI.Text = "000000000-0"
   TxtTelefonoS.Text = ""
   TxtDirS.Text = ""
   TxtEstCiv.Text = ""
   TxtNo_Dep.Text = ""
   TxtNoSoc.Text = ""
   TxtCiudadS.Text = ""
   TxtLugarTrabS.Text = ""
   TxtDirT.Text = ""
   TxtTelefonoT.Text = ""
   TxtProfesion.Text = ""
   TxtNombresC.Text = ""
   TxtApellidosC.Text = ""
   TxtLugarTrabC.Text = ""
   TxtTelefonoC.Text = ""
   TxtDirC.Text = ""
   TxtRazonSocial.Text = ""
   TxtFAX.Text = ""
   TxtCasilla.Text = ""
   TxtActividad.Text = ""
   TxtSector.Text = ""
   TxtArea.Text = ""
   If MBoxCuenta.Text <> "000000000-0" Then
   TxtNombresS.Enabled = True
   TxtApellidosS.Enabled = True
   TxtRazonSocial.Enabled = True
   CheckConyugue.Value = False
   FrameConyugue.Visible = False
   sSQL = "SELECT * FROM Cuentas " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        If .Fields("PJ") Then
            OpcN.Value = True
            PersonalJuridica True
        Else
            OpcJ.Value = True
            PersonalJuridica False
        End If
        Select Case .Fields("T")
          Case Anulado: LabelEstado.Caption = "ANULADA"
          Case Procesado: LabelEstado.Caption = "ABIERTA"
          Case Normal: LabelEstado.Caption = "NORMAL"
        End Select
        TxtNo_Dep.Text = .Fields("No_Dep")
        TxtNoSoc.Text = .Fields("No_Soc")
        CheckME.Value = .Fields("ME")
        MBoxFecha.Text = .Fields("Fecha")
    End If
   End With
   sSQL = "SELECT * FROM Clientes " _
        & "WHERE Codigo = '" & Cuenta_No & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        TxtNombresS.Text = .Fields("Nombres")
        TxtApellidosS.Text = .Fields("Apellidos")
        MBoxRUCS.Text = FormatoCodigoRUC_CI(.Fields("RUC_CI"), True)
        MBoxCI.Text = FormatoCodigoRUC_CI(.Fields("CI"), False)
        TxtTelefonoS.Text = .Fields("Telefono")
        TxtDirS.Text = .Fields("Direccion")
        TxtEstCiv.Text = .Fields("Est_Civil")
        TxtCiudadS.Text = .Fields("Ciudad")
        TxtLugarTrabS.Text = .Fields("LugarTrabajo")
        TxtDirT.Text = .Fields("DireccionT")
        TxtTelefonoT.Text = .Fields("TelefonoT")
        TxtProfesion.Text = .Fields("Profesion")
        TxtRazonSocial.Text = .Fields("Representante")
        TxtFAX.Text = .Fields("FAX")
        TxtCasilla.Text = .Fields("Casilla_Postal")
        TxtActividad.Text = .Fields("Actividad")
        TxtSector.Text = .Fields("Sector")
        TxtArea.Text = .Fields("Area")
        TxtNombresS.Enabled = False
        TxtApellidosS.Enabled = False
        TxtRazonSocial.Enabled = False
    End If
   End With
   sSQL = "SELECT * FROM Conyugue " _
        & "WHERE Cuenta_No = '" & Cuenta_No & "' "
   SelectAdodc AdoCta, sSQL
   With AdoCta.Recordset
    If .RecordCount > 0 Then
        TxtNombresC.Text = .Fields("Nombres")
        TxtApellidosC.Text = .Fields("Apellidos")
        MBoxCIC.Text = FormatoCodigoRUC_CI(.Fields("CI"), False)
        TxtLugarTrabC.Text = .Fields("LugarTrabajo")
        TxtTelefonoC.Text = .Fields("Telefono")
        TxtDirC.Text = .Fields("Direccion")
        CheckConyugue.Value = 1
        FrameConyugue.Visible = True
    Else
        CheckConyugue.Value = 0
        FrameConyugue.Visible = False
    End If
''    sSQL = "SELECT Apellidos,Nombres,T.* " _
''         & "FROM Tarjetas As T,Cuentas As C " _
''         & "WHERE T.Cuenta_No = '" & Cuenta_No & "' " _
''         & "AND T.Cuenta_No = C.Cuenta_No " _
''         & "ORDER BY Fecha_A,Tarjeta_No "
''    SelectDataGrid DGTarjeta, AdoTarjetas, sSQL
    MiFecha = PrimerDiaMes(FechaSistema)
    Dia = Day(MiFecha)
    Mes = Month(MiFecha)
    Anio = Year(MiFecha)
    If Option2.Value Then Mes = Mes - 1
    If Option3.Value Then Mes = Mes - 2
    FechaIni = Format(Dia, "00") & "/" & Format(Mes, "00") & "/" & Format(Anio, "0000")
    FechaFin = FechaSistema
    Total = 0: Saldo = 0: Contador = 1
    sSQL = "SELECT Fecha,TP,Abs(Debitos-Creditos) As Valor,Saldo_Disp,Saldo_Cont,T,Cheque,Hora " _
         & "FROM Trans_Libretas " _
         & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
         & "AND Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# and #" & BuscarFecha(FechaFin) & "# " _
         & "ORDER BY Fecha,Cuenta_No,IDT,Hora,ID "
    SelectDataGrid DGMovCta, AdoMovCta, sSQL
    With AdoMovCta.Recordset
     If .RecordCount > 0 Then
         MiFecha = .Fields("Fecha")
         Do While Not .EOF
            If MiFecha <> .Fields("Fecha") Then
               Total = Total + Saldo
               MiFecha = .Fields("Fecha")
               Contador = Contador + 1
            End If
            Saldo = .Fields("Saldo_Cont")
           .MoveNext
         Loop
         Total = Total + Saldo
         Total = Round(Total / Contador, 2)
     End If
    End With
    LabelPromedio.Caption = Format(Total, "#,##0.00")
    DGMovCta.Caption = "Cuenta No. " & Cuenta_No & Space(20) & " desde " & FechaIni & " hasta " & FechaFin
    If AdoMovCta.Recordset.RecordCount > 0 Then
       AdoMovCta.Recordset.MoveLast
    End If
    sSQL = "SELECT Fecha,Interes " _
         & "FROM Saldo_Libretas_Intereses " _
         & "WHERE Cuenta_No = '" & Cuenta_No & "' " _
         & "AND Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# and #" & BuscarFecha(FechaFin) & "# " _
         & "ORDER BY Fecha "
    SelectDataGrid DGInt, AdoIntereses, sSQL
    Total = 0
    With AdoIntereses.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Total = Total + .Fields("Interes")
           .MoveNext
         Loop
     End If
    End With
    LabelTotInt.Caption = Format(Total, "#,##0.00")
    DGInt.Caption = "Cuenta No. " & Cuenta_No
   End With
   sSQL = "SELECT TP,Credito_No,Fecha " _
        & "FROM Trans_Prestamos As P " _
        & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
        & "AND T = 'P' " _
        & "AND Cuenta_No = '" & Cuenta_No & "' " _
        & "ORDER BY TP,Fecha,Credito_No "
   SelectDataGrid DGCreditos, AdoCreditos, sSQL
   End If
End Sub

Private Sub Command3_Click()
  sSQL = "SELECT Cuenta_No FROM Cuentas " _
       & "WHERE Cuenta_No < '" & MBoxCuenta.Text & "' " _
       & "ORDER BY Cuenta_No "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
    If .RecordCount > 0 Then
       .MoveLast
        MBoxCuenta.Text = .Fields("Cuenta_No")
        ListarCuenta MBoxCuenta.Text
    End If
  End With
End Sub

Private Sub Command4_Click()
  sSQL = "SELECT Cuenta_No FROM Cuentas " _
       & "WHERE Cuenta_No > '" & MBoxCuenta.Text & "' " _
       & "ORDER BY Cuenta_No "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        MBoxCuenta.Text = .Fields("Cuenta_No")
        ListarCuenta MBoxCuenta.Text
    End If
  End With
End Sub

Private Sub CheckConyugue_Click()
  FrameConyugue.Visible = Not FrameConyugue.Visible
End Sub

Private Sub Command1_Click()
  FechaValida MBoxFecha, False
  TextoValido TxtRazonSocial, , True
  TextoValido TxtNombresS, , True
  TextoValido TxtApellidosS, , True
  TextoValido TxtProfesion, , True
  TextoValido TxtActividad, , True
  TextoValido TxtCasilla, , True
  TextoValido TxtEstCiv, , True
  TextoValido TxtNo_Dep, True, True
  TextoValido TxtCiudadS, , True
  TextoValido TxtLugarTrabS, , True
  TextoValido TxtDirS, , True
  TextoValido TxtDirT, , True
  TextoValido TxtFAX, , True
  TextoValido TxtTelefonoS, , True
  TextoValido TxtTelefonoT, , True
  TextoValido TxtArea, , True
  TextoValido TxtSector
  Mensajes = "Esta seguro de Grabar la Cuenta No. " & MBoxCuenta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "SELECT * FROM Cuentas " _
          & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
     SelectAdodc AdoCta, sSQL
     With AdoCta.Recordset
          If .RecordCount <= 0 Then
              SetAddNew AdoCta
              SetFields AdoCta, "PJ", OpcN.Value
              SetFields AdoCta, "T", Procesado
              SetFields AdoCta, "TC", "CJ"
              SetFields AdoCta, "ME", CheckME.Value
              SetFields AdoCta, "Fecha", MBoxFecha.Text
              SetFields AdoCta, "Cuenta_No", MBoxCuenta.Text
          End If
          SetFields AdoCta, "Nombres", TxtNombresS.Text
          SetFields AdoCta, "Apellidos", TxtApellidosS.Text
          SetFields AdoCta, "RUC_CI", MBoxRUCS.Text
          SetFields AdoCta, "CI", MBoxCI.Text
          SetFields AdoCta, "Telefono", TxtTelefonoS.Text
          SetFields AdoCta, "Direccion", TxtDirS.Text
          SetFields AdoCta, "Est_Civil", TxtEstCiv.Text
          SetFields AdoCta, "No_Dep", Val(TxtNo_Dep.Text)
          SetFields AdoCta, "No_Soc", Val(TxtNoSoc.Text)
          SetFields AdoCta, "Ciudad", TxtCiudadS.Text
          SetFields AdoCta, "LugarTrabajo", TxtLugarTrabS.Text
          SetFields AdoCta, "DireccionT", TxtDirT.Text
          SetFields AdoCta, "TelefonoT", TxtTelefonoT.Text
          SetFields AdoCta, "Profesion", TxtProfesion.Text
          SetFields AdoCta, "Representante", TxtRazonSocial.Text
          SetFields AdoCta, "FAX", TxtFAX.Text
          SetFields AdoCta, "Casilla_Postal", TxtCasilla.Text
          SetFields AdoCta, "Actividad", TxtActividad.Text
          SetFields AdoCta, "Sector", TxtSector.Text
          SetFields AdoCta, "Area", TxtArea.Text
          SetFields AdoCta, "CodigoU", CodigoUsuario
          SetFields AdoCta, "Item", NumEmpresa
          SetUpdate AdoCta
     End With
     If CheckConyugue.Value Then
        sSQL = "SELECT * FROM Conyugue " _
             & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
        SelectAdodc AdoCta, sSQL
        With AdoCta.Recordset
         If .RecordCount <= 0 Then
             SetAddNew AdoCta
             SetFields AdoCta, "Cuenta_No", MBoxCuenta.Text
         End If
          SetFields AdoCta, "Nombres", TxtNombresC.Text
          SetFields AdoCta, "Apellidos", TxtApellidosC.Text
          SetFields AdoCta, "CI", MBoxCIC.Text
          SetFields AdoCta, "LugarTrabajo", TxtLugarTrabC.Text
          SetFields AdoCta, "Telefono", TxtTelefonoC.Text
          SetFields AdoCta, "Direccion", TxtDirC.Text
          SetUpdate AdoCta
        End With
     End If
  End If
End Sub

Private Sub Command2_Click()
  Unload Apertura
End Sub

Private Sub Command5_Click()
 ImprimirDataApertura MBoxCuenta.Text, AdoCta, AdoAux
End Sub

Private Sub Command6_Click()
  Mensajes = "Esta seguro de desea desbloquear la Cuenta No. " & MBoxCuenta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Cuentas SET T = 'N' " _
          & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
     ConectarAdoExecute sSQL
  End If
  ListarCuenta MBoxCuenta.Text
  TxtNombresS.SetFocus
End Sub

Private Sub Command7_Click()
  FechaValida MBoxFechaT
  If MBoxTarjeta.Text <> "0000000-00000000-0" Then
  Mensajes = "Esta seguro de Grabar la Tarjeta No. " & MBoxTarjeta.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "SELECT * FROM Tarjetas " _
          & "WHERE Tarjeta_No = '" & MBoxTarjeta.Text & "' "
     SelectAdodc AdoCta, sSQL
     With AdoCta.Recordset
      If .RecordCount <= 0 Then
          SetAddNew AdoCta
          SetFields AdoCta, "TT", Procesado
          SetFields AdoCta, "Fecha_A", MBoxFechaT.Text
          SetFields AdoCta, "Cuenta_No", MBoxCuenta.Text
          SetFields AdoCta, "Tarjeta_No", MBoxTarjeta.Text
          SetUpdate AdoCta
      End If
     End With
     sSQL = "SELECT Apellidos,Nombres,T.* " _
         & "FROM Tarjetas As T,Cuentas As C " _
         & "WHERE T.Cuenta_No = '" & MBoxCuenta.Text & "' " _
         & "AND T.Cuenta_No = C.Cuenta_No " _
         & "ORDER BY Fecha_A,Tarjeta_No "
     SelectDataGrid DGTarjeta, AdoTarjetas, sSQL
  End If
  End If
End Sub

Private Sub Command8_Click()
  DBGTarjeta.Col = 3
  Tarjeta_No = DBGTarjeta.Text
  Mensajes = "Esta seguro de desea desbloquear la Tarjeta No. " & Tarjeta_No & "."
  Titulo = "Pregunta de Grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Tarjetas SET TT = 'N' " _
          & "WHERE Tarjeta_No = '" & Tarjeta_No & "' "
     ConectarAdoExecute sSQL
     sSQL = "SELECT Apellidos,Nombres,T.* " _
         & "FROM Tarjetas As T,Cuentas As C " _
         & "WHERE T.Cuenta_No = '" & MBoxCuenta.Text & "' " _
         & "AND T.Cuenta_No = C.Cuenta_No " _
         & "ORDER BY Fecha_A,Tarjeta_No "
     SelectDataGrid DGTarjeta, AdoTarjetas, sSQL
  End If
End Sub

Private Sub Command9_Click()
  DBGTarjeta.Col = 3
  Tarjeta_No = DBGTarjeta.Text
  Mensajes = "Esta seguro de desea Bloquear la Tarjeta No. " & Tarjeta_No & "."
  Titulo = "Pregunta de Grabación"
  If BoxMensaje = 6 Then
     sSQL = "UPDATE Tarjetas SET TT = 'B' " _
          & "WHERE Tarjeta_No = '" & Tarjeta_No & "' "
     ConectarAdoExecute sSQL
     sSQL = "SELECT Apellidos,Nombres,T.* " _
         & "FROM Tarjetas As T,Cuentas As C " _
         & "WHERE T.Cuenta_No = '" & MBoxCuenta.Text & "' " _
         & "AND T.Cuenta_No = C.Cuenta_No " _
         & "ORDER BY Fecha_A,Tarjeta_No "
     SelectDataGrid DGTarjeta, AdoTarjetas, sSQL
  End If
End Sub

Private Sub DGMovCta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto Apertura, AdoMovCta
End Sub

Private Sub DGTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyDelete Then
     DBGTarjeta.Col = 3
     Tarjeta_No = DBGTarjeta.Text
     Mensajes = "Esta seguro de Eliminar," & vbCrLf _
              & "la Tarjeta No. " & Tarjeta_No & "."
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = 6 Then
        sSQL = "SELECT * FROM Trans_Tarjetas " _
             & "WHERE Tarjeta_No = '" & Tarjeta_No & "' "
        SelectAdodc AdoCta, sSQL
        If AdoCta.Recordset.RecordCount <= 0 Then
           sSQL = "DELETE * FROM Tarjetas " _
                & "WHERE Tarjeta_No = '" & Tarjeta_No & "' "
           ConectarAdoExecute sSQL
        End If
        sSQL = "SELECT Apellidos,Nombres,T.* " _
             & "FROM Tarjetas As T,Cuentas As C " _
             & "WHERE T.Cuenta_No = '" & MBoxCuenta.Text & "' " _
             & "AND T.Cuenta_No = C.Cuenta_No " _
             & "ORDER BY Fecha_A,Tarjeta_No "
        SelectDataGrid DGTarjeta, AdoTarjetas, sSQL
     End If
  End If
End Sub

Private Sub Form_Activate()
   AdoCuentas.Caption = "Utilice las flechas para listar otras cuentas "
   PersonalJuridica True
   MBoxFecha.SetFocus
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm Apertura
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoMovCta
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoIntereses
End Sub

Private Sub MBoxCI_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub MBoxCIC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  'Codigo NAC
   If CtrlDown And KeyCode = vbKeyN Then
      sSQL = "SELECT Cuenta_No FROM Cuentas "
      sSQL = sSQL & "ORDER BY SUBSTRING(Cuenta_No,4,5) "
      SelectAdodc AdoAux, sSQL
      With AdoAux.Recordset
       If .RecordCount > 0 Then
          .MoveLast
           Numero = Val(Mid(.Fields("Cuenta_No"), 4, 5))
           Documento = Val(Mid(.Fields("Cuenta_No"), 10, 1))
           Documento = Documento - 1
           If Documento < 0 Then Documento = 9
           Numero = Numero + 1
           MBoxCuenta.Text = Format(Mid(.Fields("Cuenta_No"), 1, 3), "000") & Format(Numero, "00000") & "-" & Documento
       End If
      End With
   End If
End Sub

Private Sub MBoxCuenta_LostFocus()
   If MBoxCuenta.Text = "000000000-0" Then
      MBoxCuenta.Text = "123456789-1"
   End If
   ListarCuenta MBoxCuenta.Text
   If TxtNombresS.Enabled Then TxtNombresS.SetFocus Else MBoxRUCS.SetFocus
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, False
End Sub

Private Sub MBoxFechaT_GotFocus()
  MarcarTexto MBoxFechaT
End Sub

Private Sub MBoxFechaT_LostFocus()
  FechaValida MBoxFechaT
End Sub

Private Sub MBoxRUCS_KeyDown(KeyCode As Integer, Shift As Integer)
 Keys_Especiales Shift
 If ShiftDown And vbKeyHome Then
    TxtNombresS.Enabled = True
    TxtApellidosS.Enabled = True
    TxtRazonSocial.Enabled = True
 Else
    PresionoEnter KeyCode
 End If
End Sub

Private Sub MBoxRUCS_LostFocus()
''   sSQL = "SELECT T,ME,Cuenta_No," _
''        & "(Cliente & ' ' & Representante) As Propietario " _
''        & "FROM Cuentas " _
''        & "WHERE RUC_CI = '" & MBoxRUCS.Text & "' "
''   SelectDataGrid DGListCtas, AdoListCtas, sSQL
   If OpcN.Value Then TxtProfesion.SetFocus Else MBoxCI.SetFocus
End Sub


Private Sub OpcJ_Click()
PersonalJuridica False
End Sub

Private Sub OpcN_Click()
  PersonalJuridica True
End Sub

Private Sub TxtArea_GotFocus()
  TxtArea.Text = ""
End Sub

Private Sub TxtArea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtArea_LostFocus()
  TextoValido TxtArea, , True
End Sub

Private Sub TxtSector_GotFocus()
  TxtSector.Text = ""
End Sub

Private Sub TxtSector_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSector_LostFocus()
  TextoValido TxtSector
End Sub

Private Sub TxtActividad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtActividad_LostFocus()
   TextoValido TxtActividad, , True
End Sub

Private Sub TxtApellidosC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosC_LostFocus()
   TextoValido TxtApellidosC, , True
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
   TextoValido TxtApellidosS, , True
End Sub

Private Sub TxtCasilla_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtCasilla_LostFocus()
   TextoValido TxtCasilla, , True
End Sub

Private Sub TxtCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtCiudadS_LostFocus()
  TextoValido TxtCiudadS, , True
End Sub

Private Sub TxtDirC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtDirC_LostFocus()
   TextoValido TxtDirC, , True
End Sub

Private Sub TxtDirS_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtDirS_LostFocus()
   TextoValido TxtDirS, , True
End Sub

Private Sub TxtDirT_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtDirT_LostFocus()
   TextoValido TxtDirT, , True
End Sub

Private Sub TxtEstCiv_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtEstCiv_LostFocus()
   TextoValido TxtEstCiv, , True
End Sub

Private Sub TxtFAX_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtFAX_LostFocus()
   TextoValido TxtFAX, , True
End Sub

Private Sub TxtLugarTrabC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabC_LostFocus()
   TextoValido TxtLugarTrabC, , True
End Sub

Private Sub TxtLugarTrabS_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtLugarTrabS_LostFocus()
   TextoValido TxtLugarTrabS, , True
End Sub

Private Sub TxtNo_Dep_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNo_Dep_LostFocus()
   TextoValido TxtNo_Dep, True, True
End Sub

Private Sub TxtNombresC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtNombresC_LostFocus()
   TextoValido TxtNombresC, , True
End Sub

Private Sub TxtNombresS_KeyDown(KeyCode As Integer, Shift As Integer)
 PresionoEnter KeyCode
End Sub

Private Sub TxtNombresS_LostFocus()
  TextoValido TxtNombresS, , True
End Sub

Private Sub TxtNoSoc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNoSoc_LostFocus()
  TextoValido TxtNoSoc, True
  If Val(TxtNoSoc.Text) <= 0 Then TxtNoSoc.Text = "1"
End Sub

Private Sub TxtProfesion_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtProfesion_LostFocus()
  TextoValido TxtProfesion, , True
End Sub

Private Sub TxtRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRazonSocial_LostFocus()
   TextoValido TxtRazonSocial, , True
End Sub

Private Sub TxtTelefonoC_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoC_LostFocus()
   TextoValido TxtTelefonoC, , True
End Sub

Private Sub TxtTelefonoS_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoS_LostFocus()
  TextoValido TxtTelefonoS, , True
End Sub

Private Sub TxtTelefonoT_KeyDown(KeyCode As Integer, Shift As Integer)
PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoT_LostFocus()
   TextoValido TxtTelefonoT, , True
End Sub

