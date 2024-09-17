VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FClientes 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Apertura de cuenta"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   14715
   WindowState     =   1  'Minimized
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   93
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
         NumButtons      =   24
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
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Crea un nuevo Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modifica el Registro"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar el Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bloquear"
            Object.ToolTipText     =   "Bloquear el Registro"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Activa el Registro Bloqueado"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar los Datos del Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxC"
            Object.ToolTipText     =   "Asignar a Cuentas por Cobrar en Contabilidad"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "CxP"
            Object.ToolTipText     =   "Asignar a Cuentas por Pagar en Contabilidad"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ahorros"
            Object.ToolTipText     =   "Asignar Cuenta de Ahorros"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Tarjetas"
            Object.ToolTipText     =   "Asigna Tarjeta de Débito"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RolPago"
            Object.ToolTipText     =   "Asignar a Rol de Pago"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RenumFactMult"
            Object.ToolTipText     =   "Renumerar Clientes de Facturación sin C.I./RUC"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Suscripcion"
            Object.ToolTipText     =   "Ingresar suscripción"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Facturacion"
            Object.ToolTipText     =   "Asignar Cliente de Facturación"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Emails"
            Object.ToolTipText     =   "Genera la lista de Mails en un archivo de texto"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnFacMult"
            Object.ToolTipText     =   "Desabilitar Asignacion de Facturacion"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Archivo_Excel"
            Object.ToolTipText     =   "Envia la Base a Excel"
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Generar_Educativo"
            Object.ToolTipText     =   "Generar nomica estudiantes del Educativo"
            Object.Tag             =   ""
            ImageIndex      =   24
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cambio_Ejecutivo"
            Object.ToolTipText     =   "Cambia Ejecutivo de Ventas de Clientes"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Activar_CE"
            Object.ToolTipText     =   "Activar Usuarios a Comprobantes Electronicos"
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7770
      Picture         =   "Clientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   735
      Width           =   330
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
      Left            =   8190
      MaxLength       =   13
      TabIndex        =   3
      ToolTipText     =   "<Alt+F2> Codigo Automático"
      Top             =   1365
      Width           =   1590
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9765
      Picture         =   "Clientes.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   1155
      Width           =   645
   End
   Begin VB.TextBox TxtDescuento 
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
      Left            =   11760
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "Clientes.frx":0DA0
      Top             =   2625
      Width           =   1170
   End
   Begin VB.Frame FrmPatronBusqueda 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PATRON DE BUSQUEDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   13020
      TabIndex        =   88
      Top             =   3570
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ListBox LstCampos 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   105
         TabIndex        =   89
         Top             =   210
         Width           =   4005
      End
      Begin VB.TextBox TxtCIRUC 
         BackColor       =   &H00C00000&
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
         TabIndex        =   90
         Top             =   2730
         Width           =   4005
      End
   End
   Begin VB.TextBox TxtContacto 
      BeginProperty Font 
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
      MaxLength       =   60
      TabIndex        =   15
      Top             =   2625
      Width           =   3585
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
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1995
      Width           =   1485
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
      Left            =   9660
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1995
      Width           =   1590
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
      Left            =   11235
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1995
      Width           =   1695
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
      Left            =   11655
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1365
      Width           =   1275
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
      Left            =   10395
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1365
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      Left            =   630
      Picture         =   "Clientes.frx":0DA4
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   8295
      Width           =   330
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Clientes.frx":143A
      DataSource      =   "AdoCliente"
      Height          =   1935
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "<Ctrl+B> Buscar Datos"
      Top             =   1050
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   3413
      _Version        =   393216
      Style           =   1
      ForeColor       =   8388608
      Text            =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoCreditos 
      Height          =   330
      Left            =   210
      Top             =   1890
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   11235
      TabIndex        =   84
      Top             =   8715
      Width           =   750
   End
   Begin VB.ComboBox CTipoPersona 
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
      Height          =   315
      Left            =   9765
      TabIndex        =   87
      Text            =   "Combo1"
      Top             =   735
      Width           =   3165
   End
   Begin VB.CommandButton Command4 
      Height          =   330
      Left            =   210
      Picture         =   "Clientes.frx":1453
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   8295
      Width           =   330
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
      Left            =   12075
      TabIndex        =   85
      Top             =   8715
      Width           =   750
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5685
      Left            =   105
      TabIndex        =   22
      Top             =   3465
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   10028
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "&1.- Datos Principales"
      TabPicture(0)   =   "Clientes.frx":1AE9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label25"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label31"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label34"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label35"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label26"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label22"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label36"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label37"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label38"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label15"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label16"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label32"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label18"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label28"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label33"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "MBFechaN"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "MBFecha"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtActividad"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtRazonSocial"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtLugarTrabS"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtCredito"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtComision"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TxtNumero"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TxtEmail"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "CCiudadS"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TxtCasilla"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "CNacion"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "CProvincia"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtProfesion"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtDirS"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TxtNo_Dep"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TxtDirT"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "CEstado"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TxtLDirs"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "LstProductos"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtPlan"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtCalificacion"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TxtEmail2"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "CTipoProv"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "CParteR"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "CMedidor"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtCodigo"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TxtApellidosS"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "OpcF"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "OpcM"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "CheqRISE"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "CheqContEsp"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "CheqDr"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).ControlCount=   60
      TabCaption(1)   =   "&2.- Ordenados por RUC/CI/Cod. Banco"
      TabPicture(1)   =   "Clientes.frx":1B05
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGCliente"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox CheqDr 
         Caption         =   "Dr."
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
         Left            =   12075
         TabIndex        =   92
         Top             =   630
         Width           =   645
      End
      Begin VB.CheckBox CheqContEsp 
         Caption         =   "Cont. Esp."
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
         Left            =   10920
         TabIndex        =   31
         Top             =   420
         Width           =   1275
      End
      Begin VB.CheckBox CheqRISE 
         Caption         =   "RISE"
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
         Left            =   10920
         TabIndex        =   32
         Top             =   630
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
         Left            =   9555
         TabIndex        =   29
         Top             =   420
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
         Left            =   9555
         TabIndex        =   30
         Top             =   630
         Width           =   1275
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
         MaxLength       =   180
         TabIndex        =   25
         Top             =   645
         Width           =   6315
      End
      Begin VB.TextBox TxtCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   60
         TabIndex        =   24
         Top             =   420
         Width           =   1380
      End
      Begin VB.ComboBox CMedidor 
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
         Left            =   11445
         TabIndex        =   60
         Text            =   "Soltero"
         Top             =   2205
         Width           =   1275
      End
      Begin VB.ComboBox CParteR 
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
         Left            =   8715
         TabIndex        =   28
         Top             =   645
         Width           =   750
      End
      Begin VB.ComboBox CTipoProv 
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
         TabIndex        =   27
         Top             =   645
         Width           =   2325
      End
      Begin VB.TextBox TxtEmail2 
         BeginProperty Font 
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
         Top             =   2730
         Width           =   6315
      End
      Begin VB.TextBox TxtCalificacion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11445
         MaxLength       =   2
         TabIndex        =   76
         Top             =   3255
         Width           =   1275
      End
      Begin VB.TextBox TxtPlan 
         BeginProperty Font 
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
         MaxLength       =   30
         TabIndex        =   64
         Top             =   2745
         Width           =   2115
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
         Height          =   1230
         Left            =   8820
         TabIndex        =   80
         Top             =   3900
         Width           =   3900
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
         Height          =   1275
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   78
         Text            =   "Clientes.frx":1B21
         Top             =   3900
         Width           =   8625
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
         TabIndex        =   52
         Text            =   "Soltero"
         Top             =   2220
         Width           =   1170
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
         Left            =   6405
         MaxLength       =   50
         TabIndex        =   74
         Top             =   3255
         Width           =   5055
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
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   54
         Top             =   2220
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
         TabIndex        =   34
         Top             =   1170
         Width           =   6315
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
         MaxLength       =   50
         TabIndex        =   70
         Top             =   3270
         Width           =   3585
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
         Left            =   2520
         TabIndex        =   42
         Text            =   "PICHINCHA"
         Top             =   1695
         Width           =   3900
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
         TabIndex        =   40
         Text            =   "ECUADOR"
         Top             =   1680
         Width           =   2430
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
         Left            =   8505
         MaxLength       =   15
         TabIndex        =   56
         Top             =   2220
         Width           =   1695
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
         Left            =   6405
         TabIndex        =   44
         Top             =   1695
         Width           =   3795
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
         Left            =   7770
         MaxLength       =   50
         TabIndex        =   38
         Top             =   1155
         Width           =   4950
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
         Left            =   6405
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1170
         Width           =   1380
      End
      Begin VB.TextBox TxtComision 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10185
         MaxLength       =   3
         TabIndex        =   58
         Top             =   2220
         Width           =   1275
      End
      Begin VB.TextBox TxtCredito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11445
         MaxLength       =   3
         TabIndex        =   68
         Top             =   2745
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
         Left            =   3675
         MaxLength       =   30
         TabIndex        =   72
         Top             =   3270
         Width           =   2745
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
         MaxLength       =   120
         TabIndex        =   50
         Top             =   2220
         Width           =   6315
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
         Left            =   8505
         MaxLength       =   25
         TabIndex        =   66
         Top             =   2730
         Width           =   2955
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   10185
         TabIndex        =   46
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1680
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
      Begin MSMask.MaskEdBox MBFechaN 
         Height          =   330
         Left            =   11445
         TabIndex        =   48
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   1695
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
      Begin MSDataGridLib.DataGrid DGCliente 
         Bindings        =   "Clientes.frx":1B25
         Height          =   4740
         Left            =   -74895
         TabIndex        =   91
         Top             =   435
         Width           =   10830
         _ExtentX        =   19103
         _ExtentY        =   8361
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
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Medidor No."
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
         Left            =   11445
         TabIndex        =   59
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL2 (CORREO ELECTRONICO 2)"
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
         Top             =   2520
         Width           =   6315
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL (CORREO ELECTRONICO)"
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
         Left            =   7770
         TabIndex        =   37
         Top             =   945
         Width           =   4950
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
         Left            =   105
         TabIndex        =   23
         Top             =   435
         Width           =   4950
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* TIPO PROV. Y PARTE RELAC."
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
         TabIndex        =   26
         Top             =   435
         Width           =   3060
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CALIFICA."
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
         Left            =   11445
         TabIndex        =   75
         Top             =   3045
         Width           =   1275
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PLAN AFILIACION"
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
         TabIndex        =   63
         Top             =   2535
         Width           =   2115
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
         Left            =   6405
         TabIndex        =   35
         Top             =   945
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
         Height          =   225
         Left            =   8295
         TabIndex        =   83
         Top             =   5370
         Width           =   2535
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
         Left            =   8820
         TabIndex        =   79
         Top             =   3690
         Width           =   3900
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
         TabIndex        =   77
         Top             =   3690
         Width           =   8625
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
         Left            =   11445
         TabIndex        =   47
         Top             =   1485
         Width           =   1275
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
         Left            =   6405
         TabIndex        =   73
         Top             =   3045
         Width           =   5055
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
         Left            =   6405
         TabIndex        =   43
         Top             =   1485
         Width           =   3795
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
         TabIndex        =   33
         Top             =   960
         Width           =   6315
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
         TabIndex        =   69
         Top             =   3060
         Width           =   3585
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
         Left            =   2520
         TabIndex        =   41
         Top             =   1485
         Width           =   3900
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
         TabIndex        =   39
         Top             =   1485
         Width           =   2430
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
         TabIndex        =   51
         Top             =   2010
         Width           =   1170
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " COMISION"
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
         TabIndex        =   57
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CREDITOS"
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
         Left            =   11445
         TabIndex        =   67
         Top             =   2535
         Width           =   1275
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
         Left            =   3675
         TabIndex        =   71
         Top             =   3060
         Width           =   2745
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
         Left            =   7560
         TabIndex        =   53
         Top             =   2010
         Width           =   960
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
         Left            =   8505
         TabIndex        =   55
         Top             =   2010
         Width           =   1695
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
         TabIndex        =   49
         Top             =   2010
         Width           =   6315
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ACTIVIDAD (CTE/AHO)"
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
         Left            =   8505
         TabIndex        =   65
         Top             =   2520
         Width           =   2955
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
         Left            =   10185
         TabIndex        =   45
         Top             =   1470
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc AdoTarjetas 
      Height          =   330
      Left            =   2730
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
      Left            =   2730
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
      Left            =   2730
      Top             =   2205
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
      Left            =   210
      Top             =   1260
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2730
      Top             =   1890
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
   Begin MSAdodcLib.Adodc AdoBuses 
      Height          =   330
      Left            =   210
      Top             =   2205
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Buses"
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
      Left            =   210
      Top             =   2520
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSMask.MaskEdBox MBCtaxCob 
      Height          =   330
      Left            =   11235
      TabIndex        =   21
      Top             =   3045
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc AdoEjec 
      Height          =   330
      Left            =   210
      Top             =   1575
      Width           =   2325
      _ExtentX        =   4101
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
      Caption         =   "Ejec"
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
   Begin MSDataListLib.DataCombo DCEjec 
      Bindings        =   "Clientes.frx":1B3E
      DataSource      =   "AdoEjec"
      Height          =   315
      Left            =   1995
      TabIndex        =   19
      Top             =   3045
      Visible         =   0   'False
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin VB.Label LblSRI 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   13020
      TabIndex        =   95
      Top             =   735
      Width           =   6105
   End
   Begin VB.Label Label39 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ejecutivo de Venta"
      BeginProperty Font 
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
      Top             =   3045
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Asignar Cuenta por Cobrar"
      BeginProperty Font 
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
      TabIndex        =   20
      Top             =   3045
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Descuento"
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
      Left            =   11760
      TabIndex        =   16
      Top             =   2415
      Width           =   1170
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONTACTO"
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
      Left            =   8190
      TabIndex        =   14
      Top             =   2415
      Width           =   3585
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
      Left            =   8190
      TabIndex        =   8
      Top             =   1785
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
      Left            =   9660
      TabIndex        =   10
      Top             =   1785
      Width           =   1590
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   11235
      TabIndex        =   12
      Top             =   1785
      Width           =   1695
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
      Left            =   11655
      TabIndex        =   6
      Top             =   1155
      Width           =   1275
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
      Left            =   10395
      TabIndex        =   4
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
      Left            =   8190
      TabIndex        =   2
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO PERSONA"
      BeginProperty Font 
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
      TabIndex        =   86
      Top             =   735
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
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
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   7680
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   13125
      Top             =   7980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   26
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":2188
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":24A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":27BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":2AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":2DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":310A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":3424
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":373E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":3A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":3D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":12B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":12E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":13138
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":13452
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":13604
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":13E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1405C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":14376
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":14690
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":149AA
            Key             =   "EXCELL"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":14B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1CFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1D30C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Clientes.frx":1D626
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
Option Explicit

Dim Actualiza_Buses As Boolean
Dim Campos_Clientes As String
Dim Temp_TD As String
Dim Temp_CI_RUC As String
Dim Temp_Cliente As String
Dim Temp_Ejecutivo As String
Dim Temp_Plan As String

Public Sub ListarEmailsClientes()
Dim NumFile As Integer
Dim NumFile1 As Integer
Dim TotalCampo As Integer

Dim ContadorReg As Long

Dim RutaGeneraFile As String
Dim CaptionOld As String
Dim NombreFile As String
Dim CadFileReg As String
Dim ValorBool As String

RatonReloj
ContadorReg = 0
If FileResp <= 0 Then FileResp = 1
sSQL = "SELECT Codigo,Cliente,Email,Email2 " _
     & "FROM Clientes " _
     & "WHERE Codigo <> '.' " _
     & "ORDER BY Cliente "
Select_Adodc AdoListCtas, sSQL
With AdoListCtas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TotalReg = .RecordCount
     TotalCampo = .fields.Count - 1
     NombreFile = "Email " & Format$(Day(FechaSistema), "00") & "-" & Format$(Month(FechaSistema), "00") _
                & "-" & Format$(Year(FechaSistema), "0000") & ".txt"
     RutaGeneraFile = LeftStrg(CurDir$, 2) & "\SYSBASES\EMAILS\" & NombreFile
     NumFile = FreeFile
     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
     
     NombreFile = "Email Clientes " & Format$(Day(FechaSistema), "00") & "-" & Format$(Month(FechaSistema), "00") _
                & "-" & Format$(Year(FechaSistema), "0000") & ".csv"
     RutaGeneraFile = LeftStrg(CurDir$, 2) & "\SYSBASES\EMAILS\" & NombreFile
     NumFile1 = FreeFile
     Open RutaGeneraFile For Output As #NumFile1 ' Abre el archivo.
     Print #NumFile1, "CLIENTES;CORREOS ELECTRONICOS"
     Do While Not .EOF
        ContadorReg = ContadorReg + 1
        FClientes.Caption = "Creando Archivo de Mails: " _
                          & "(" & Format$(ContadorReg / TotalReg, "00%") _
                          & ") " & String(ContadorReg Mod 40, "|")
        ValorBool = False
        CaptionOld = ""
        If EsUnEmail(.fields("Email")) Then
           CaptionOld = .fields("Email") & ";"
           ValorBool = True
        End If
        If EsUnEmail(.fields("Email2")) Then
           CaptionOld = CaptionOld & .fields("Email2") & ";"
           ValorBool = True
        End If
        If ValorBool Then
           Print #NumFile, CaptionOld
           Print #NumFile1, .fields("Cliente") & ";" & CaptionOld
        End If
       .MoveNext
     Loop
     Close #NumFile
     Close #NumFile1
 End If
End With
RatonNormal
FClientes.Caption = "APERTURA DE CLIENTES"
MsgBox "Se ha procesado un archivo en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Public Sub GrabarCliente()
  Si_No = False
  Nuevo = False
  FechaValida MBFecha
  FechaValida MBFechaN
  TextoValido TxtGrupo, , True
  TextoValido TxtCI_RUC, , True
  TextoValido TxtRazonSocial, , True
  TextoValido TxtApellidosS, , True
  TextoValido TxtProfesion, , True
  TextoValido TxtActividad, , True
  TextoValido TxtEmail
  TextoValido TxtEmail2
  TextoValido TxtContacto, , True
  TextoValido TxtCasilla, , True
  TextoValido TxtNo_Dep, True, True
  TextoValido TxtLugarTrabS, , True
  TextoValido TxtDirS, , True
  TextoValido TxtPlan, , True
  TextoValido TxtDirT, , True
  TextoValido TxtFAX, , True
  TextoValido TxtTelefonoS, , True
  TextoValido TxtTelefonoT, , True
  TextoValido TxtNumero, , True
  TextoValido TxtComision, True
  TextoValido TxtCredito, True, , 0
  TextoValido TxtCalificacion, , True
  TextoValido TxtDescuento, True, , 2
    
  If TxtCodigo = "9999999999" Then
     TxtApellidosS = "CONSUMIDOR FINAL"
     TxtCI_RUC = "9999999999999"
     TipoBenef = "R"
  End If
  If TxtCI_RUC = Ninguno Then
     MsgBox "No se puede grabar, C.I./R.U.C. deben tener valores"
  Else
     Mensajes = "Esta seguro de Grabar"
     Titulo = "Pregunta de Grabación"
     If BoxMensaje = vbYes Then
        CodigoEjecutivo = Ninguno
        If DCEjec.Text = "" Then DCEjec.Text = Ninguno
        If AdoEjec.Recordset.RecordCount > 0 Then
           AdoEjec.Recordset.MoveFirst
           AdoEjec.Recordset.Find ("Ejecutivo = '" & DCEjec.Text & "' ")
           If Not AdoEjec.Recordset.EOF Then CodigoEjecutivo = AdoEjec.Recordset.fields("Codigo")
        End If
        
        sSQL = "SELECT " & Full_Fields("Clientes") & " " _
             & "FROM Clientes " _
             & "WHERE Codigo = '" & TxtCodigo & "' "
        Select_Adodc AdoListCtas, sSQL
        With AdoListCtas.Recordset
         If .RecordCount > 0 Then
             Codigo = TxtCodigo
             If AdoListCtas.Recordset.fields("Cliente") <> TxtApellidosS Then Nuevo = True
         Else
             Nuevo = True
             Codigo = Tipo_RUC_CI.Codigo_RUC_CI
             T = Normal
             SetAddNew AdoListCtas
             SetFields AdoListCtas, "T", T
             SetFields AdoListCtas, "Codigo", Codigo
         End If
        End With
        RatonReloj
        
        SetFields AdoListCtas, "CI_RUC", TxtCI_RUC
        SetFields AdoListCtas, "TD", TipoBenef
        SetFields AdoListCtas, "Fecha", MBFecha
        SetFields AdoListCtas, "Fecha_N", MBFechaN
        SetFields AdoListCtas, "Cliente", TxtApellidosS
        SetFields AdoListCtas, "Telefono", TxtTelefonoS
        SetFields AdoListCtas, "Celular", TxtCelular
        SetFields AdoListCtas, "Email", TxtEmail
        SetFields AdoListCtas, "Email2", TxtEmail2
        SetFields AdoListCtas, "Grupo", TxtGrupo
        SetFields AdoListCtas, "Direccion", TxtDirS
        SetFields AdoListCtas, "DirNumero", TxtNumero
        SetFields AdoListCtas, "No_Dep", Val(TxtNo_Dep)
        SetFields AdoListCtas, "Lugar_Trabajo", TxtLugarTrabS
        SetFields AdoListCtas, "DireccionT", TxtDirT
        SetFields AdoListCtas, "TelefonoT", TxtTelefonoT
        SetFields AdoListCtas, "Profesion", TxtProfesion
        SetFields AdoListCtas, "Representante", TxtRazonSocial
        SetFields AdoListCtas, "FAX", TxtFAX
        SetFields AdoListCtas, "Casilla", TxtCasilla
        SetFields AdoListCtas, "Actividad", TxtActividad
        SetFields AdoListCtas, "Est_Civil", MidStrg(CEstado, 1, 1)
        SetFields AdoListCtas, "Ciudad", DN.Descripcion
        SetFields AdoListCtas, "Prov", DN.CProvincia
        SetFields AdoListCtas, "Pais", DN.CPais
        SetFields AdoListCtas, "CodigoU", CodigoUsuario
        SetFields AdoListCtas, "Porc_C", TxtComision.Text
        SetFields AdoListCtas, "Credito", TxtCredito.Text
        SetFields AdoListCtas, "Calificacion", TxtCalificacion
        SetFields AdoListCtas, "Plan_Afiliado", TxtPlan
        SetFields AdoListCtas, "Contacto", TxtContacto
        SetFields AdoListCtas, "Cta_CxP", CambioCodigoCta(MBCtaxCob)
        SetFields AdoListCtas, "Cod_Ejec", CodigoEjecutivo
        If CheqContEsp.value = 0 Then
           SetFields AdoListCtas, "Especial", False
        Else
           SetFields AdoListCtas, "Especial", True
        End If
        If CheqRISE.value = 0 Then
           SetFields AdoListCtas, "RISE", False
        Else
           SetFields AdoListCtas, "RISE", True
        End If
        If OpcM.value Then
           SetFields AdoListCtas, "Sexo", "M"
        Else
           SetFields AdoListCtas, "Sexo", "F"
        End If
        If CheqDr.value Then
           SetFields AdoListCtas, "Asignar_Dr", True
        Else
           SetFields AdoListCtas, "Asignar_Dr", False
        End If
        If Modulo = "FACTURACION" Then
           SetFields AdoListCtas, "FA", adTrue
        Else
           SetFields AdoListCtas, "FA", adFalse
        End If
        Select Case CTipoProv
           Case "OTRO"
                SetFields AdoListCtas, "Tipo_Pasaporte", "00"
           Case "PERSONA NATURAL"
                SetFields AdoListCtas, "Tipo_Pasaporte", "01"
           Case Else
                SetFields AdoListCtas, "Tipo_Pasaporte", "02"
        End Select
        If CParteR.Text = "SI" Then
           SetFields AdoListCtas, "Parte_Relacionada", "SI"
        Else
           SetFields AdoListCtas, "Parte_Relacionada", "NO"
        End If
        SetUpdate AdoListCtas
     End If
  End If
 'Ingresamos el historial de direcciones
  Si_No = False
  sSQL = "SELECT TOP 1 * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Dato = 'DIRECCION' " _
       & "ORDER BY Fecha_Registro DESC "
  Select_Adodc AdoCta, sSQL
  With AdoCta.Recordset
   If .RecordCount > 0 Then
       If .fields("Lugar_Trabajo") <> TxtLugarTrabS.Text Then Si_No = True
       If .fields("Direccion") <> TxtDirS.Text Then Si_No = True
       If .fields("DireccionT") <> TxtDirT.Text Then Si_No = True
       If .fields("TelefonoT") <> TxtTelefonoT.Text Then Si_No = True
       If .fields("Telefono") <> TxtTelefonoS.Text Then Si_No = True
       If .fields("Celular") <> TxtCelular.Text Then Si_No = True
       If .fields("FAX") <> TxtFAX.Text Then Si_No = True
       If .fields("Ciudad") <> CCiudadS.Text Then Si_No = True
      'If .Fields("Prov") <> SinEspaciosIzq(CProvincia.Text) Then Si_No = True
      'If .Fields("Pais") <> SinEspaciosIzq(CNacion.Text) Then Si_No = True
       If .fields("Descuento") <> CCur(TxtDescuento) Then Si_No = True
   Else
       Si_No = True
   End If
  End With
  If Si_No Then
     SetAddNew AdoCta
     SetFields AdoCta, "Fecha_Registro", FechaSistema
     SetFields AdoCta, "Codigo", Codigo
     SetFields AdoCta, "Lugar_Trabajo", TxtLugarTrabS
     SetFields AdoCta, "Direccion", TxtDirS
     SetFields AdoCta, "DireccionT", TxtDirT
     SetFields AdoCta, "TelefonoT", TxtTelefonoT
     SetFields AdoCta, "Telefono", TxtTelefonoS
     SetFields AdoCta, "Celular", TxtCelular
     SetFields AdoCta, "FAX", TxtFAX
     SetFields AdoCta, "Ciudad", CCiudadS
     SetFields AdoCta, "Prov", SinEspaciosIzq(CProvincia)
     SetFields AdoCta, "Pais", SinEspaciosIzq(CNacion)
     SetFields AdoCta, "CodigoU", CodigoUsuario
     SetFields AdoCta, "Descuento", CCur(TxtDescuento)
     SetFields AdoCta, "Item", NumEmpresa
     SetFields AdoCta, "Tipo_Dato", "DIRECCION"
     SetUpdate AdoCta
  End If
 'Actualizar_Bus Codigo, TxtPlan
  If Len(TxtPlan) > 3 And MidStrg(TxtPlan, 1, 3) = "BUS" Then
     Mensajes = "Desea actualizar los pagos"
     Titulo = "Pregunta de Actualizacion"
     If BoxMensaje = vbYes Then Actualizar_Bus Codigo, TxtPlan, TxtGrupo
  End If
 
 'Cambio de Ejecutivo
  If Temp_Ejecutivo <> DCEjec.Text Then
     CodigoVen = CodigoEjecutivo
     If AdoEjec.Recordset.RecordCount > 0 Then
        AdoEjec.Recordset.MoveFirst
        AdoEjec.Recordset.Find ("Ejecutivo = '" & Temp_Ejecutivo & "' ")
        If Not AdoEjec.Recordset.EOF Then CodigoEjecutivo = AdoEjec.Recordset.fields("Codigo")
     End If
     Codigo1 = CambioCodigoCta(MBCtaxCob)
       sSQL = "UPDATE Facturas " _
            & "SET Cod_Ejec = '" & CodigoVen & "' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Cod_Ejec = '" & CodigoEjecutivo & "' " _
            & "AND Cta_CxP = '" & Codigo1 & "' "
       Ejecutar_SQL_SP sSQL
       
       If SQL_Server Then
          sSQL = "UPDATE Detalle_Factura " _
               & "SET Cod_Ejec = F.Cod_Ejec " _
               & "FROM Detalle_Factura As DF, Facturas As F "
       Else
          sSQL = "UPDATE Detalle_Factura As DF, Facturas As F " _
               & "SET DF.Cod_Ejec = F.Cod_Ejec "
       End If
       sSQL = sSQL _
            & "WHERE F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND F.Cta_CxP = '" & Codigo1 & "' " _
            & "AND DF.Item = F.Item " _
            & "AND DF.Periodo = F.Periodo " _
            & "AND DF.TC = F.TC " _
            & "AND DF.Serie = F.Serie " _
            & "AND DF.Factura = F.Factura " _
            & "AND DF.Autorizacion = F.Autorizacion "
       Ejecutar_SQL_SP sSQL
       
       If SQL_Server Then
          sSQL = "UPDATE Trans_Abonos " _
               & "SET Cod_Ejec = F.Cod_Ejec " _
               & "FROM Trans_Abonos As DF, Facturas As F "
       Else
          sSQL = "UPDATE Trans_Abonos As DF, Facturas As F " _
               & "SET DF.Cod_Ejec = F.Cod_Ejec "
       End If
       sSQL = sSQL _
            & "WHERE F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND F.Cta_CxP = '" & Codigo1 & "' " _
            & "AND DF.Item = F.Item " _
            & "AND DF.Periodo = F.Periodo " _
            & "AND DF.TP = F.TC " _
            & "AND DF.Serie = F.Serie " _
            & "AND DF.Factura = F.Factura " _
            & "AND DF.Autorizacion = F.Autorizacion "
       Ejecutar_SQL_SP sSQL
      
       sSQL = "SELECT C.Cta_CxP, CC.Cuenta, COUNT(C.Codigo) " _
            & "FROM Catalogo_Cuentas As CC, Clientes As C " _
            & "WHERE CC.Item = '" & NumEmpresa & "' " _
            & "AND CC.Periodo = '" & Periodo_Contable & "' " _
            & "AND CC.Codigo = C.Cta_CxP " _
            & "GROUP BY C.Cta_CxP, CC.Cuenta " _
            & "ORDER BY CC.Cuenta "
       Select_Adodc AdoAux, sSQL
       With AdoAux.Recordset
        If .RecordCount > 0 Then
            Do While Not .EOF
               Codigo1 = .fields("Cta_CxP")
               Codigo2 = .fields("Cuenta")
               I = InStr(Codigo2, "ZONA")
              ' MsgBox Codigo2
               If I > 0 Then
                  Codigo3 = TrimStrg(MidStrg(Codigo2, I, 10))
                  sSQL = "UPDATE Clientes " _
                       & "SET Grupo = '" & Codigo3 & "' " _
                       & "WHERE Cta_CxP = '" & Codigo1 & "' "
                  Ejecutar_SQL_SP sSQL
               End If
              .MoveNext
            Loop
        End If
       End With
  End If
  If Nuevo Then
     Control_Procesos Normal, "Nuevo Beneficiario: " & TxtApellidosS
     ListarClientes
  Else
     Control_Procesos Normal, "Grabar/Modificar: " & TxtApellidosS
  End If
  MsgBox "ACTUALIZACION EXITOSA"
  DCCliente.SetFocus
End Sub

Public Sub DatosNuevos()
   FechaValida MBFecha
   FechaValida MBFechaN
   TxtCodigo = "Ninguno"
   TxtApellidosS = ""
   TxtCI_RUC.Text = ""
   TxtTelefonoS.Text = ""
   TxtDirS.Text = ""
   TxtNo_Dep.Text = ""
   TxtGrupo.Text = NumEmpresa
   CCiudadS.Text = ""
   TxtLugarTrabS.Text = ""
   TxtDirT.Text = ""
   TxtTelefonoT.Text = ""
   TxtProfesion.Text = ""
   'TxtNombresC.Text = ""
   'TxtLugarTrabC.Text = ""
   'TxtTelefonoC.Text = ""
   TxtComision.Text = "0"
   TxtRazonSocial.Text = ""
   TxtFAX.Text = ""
   TxtCasilla.Text = ""
   TxtActividad.Text = ""
   TxtEmail.Text = ""
   TxtEmail2.Text = ""
   LstProductos.Clear
   TxtLDirs.Text = ""
   TxtNumero.Text = ""
   MBFecha = FechaSistema
   TxtApellidosS.Enabled = True
   TxtRazonSocial.Enabled = True
   CheqContEsp.value = 0
   CheqRISE.value = 0
   CodigoEjecutivo = Ninguno
End Sub

Public Sub ListarClientes(Optional Campo As String, Optional Busqueda As String)
Dim TipoPersona As String

  If CTipoPersona.Text <> "Todos" Then Actualizar_Tipo_Clientes_SP
  
 'Listamos los clients segun el Tipo
  Campos_Clientes = "SELECT SUBSTRING(Codigo,1,3) Tipo_Cont, T, TD, CI_RUC, Cliente, Contacto, Grupo, Prov, Ciudad, Direccion, Fecha_N, DireccionT, Telefono, " _
                  & "Celular, TelefonoT, FAX, Sexo, RISE, Especial, Email, Email2, EmailR, Cta_CxP, Cod_Ejec, Codigo " _
                  & "FROM Clientes "
  sSQL = Campos_Clientes
  Select Case CTipoPersona
    Case "R.U.C."
         sSQL = sSQL & "WHERE TD = 'R' "
    Case "Cedulas"
         sSQL = sSQL & "WHERE TD = 'C' "
    Case "Pasaporte"
         sSQL = sSQL & "WHERE TD = 'P' "
    Case "Clientes"
         sSQL = sSQL & "WHERE FA <> " & Val(adFalse) & " "
    Case "Clientes Varones"
         sSQL = sSQL _
              & "WHERE FA <> " & Val(adFalse) & " " _
              & "AND Sexo = 'M' "
    Case "Clientes Mujeres"
         sSQL = sSQL _
              & "WHERE FA <> " & Val(adFalse) & " " _
              & "AND Sexo = 'F' "
    Case "RISE"
         sSQL = sSQL & "WHERE RISE <> " & Val(adFalse) & " "
    Case "Contribuyente Especial"
         sSQL = sSQL & "WHERE Especial <> " & Val(adFalse) & " "
    Case "Empleados"
         sSQL = sSQL & "WHERE Tipo_Cliente LIKE '%ROL,%' "
    Case "Cuentas por Cobrar"
         sSQL = sSQL & "WHERE Tipo_Cliente LIKE '%CXC,%' "
    Case "Cuentas por Pagar"
         sSQL = sSQL & "WHERE Tipo_Cliente LIKE '%CXP,%' "
    Case "Libretas de Ahorro"
         sSQL = sSQL & "WHERE Tipo_Cliente LIKE '%AHR,%' "
    Case "Clientes Debitos"
         sSQL = sSQL _
              & "WHERE MidStrg(Actividad,1,5) = 'TRANS' " _
              & "AND FA <> " & Val(adFalse) & " "
    Case "Clientes Transaferencias"
         sSQL = sSQL _
              & "WHERE MidStrg(Actividad,1,5) = 'TRANS' " _
              & "AND FA = " & Val(adFalse) & " "
    Case "Clientes Descuentos"
         sSQL = sSQL _
              & "WHERE Tipo_Cliente LIKE '%DESC,%' " _
              & "AND Tipo_Cliente LIKE '%" & NumEmpresa & "%' "
    Case "Clientes Sin Email"
         sSQL = sSQL _
              & "WHERE FA <> " & Val(adFalse) & " " _
              & "AND LEN(Email+Email2+EmailR) <= 4 "
    Case "Clientes Sin Representantes"
         sSQL = sSQL _
              & "WHERE FA <> " & Val(adFalse) & " " _
              & "AND LEN(Representante) <= 1 " _
              & "AND LEN(CI_RUC_R) <= 1 "
    Case Else
         If Len(Campo) > 1 And Len(Busqueda) >= 1 Then
            sSQL = Campos_Clientes
            Select Case Campo
              Case "CI_RUC_Representante"
                   sSQL = sSQL _
                        & "WHERE CI_RUC_R LIKE '%" & Busqueda & "%' "
              Case "Representante"
                   sSQL = sSQL _
                        & "WHERE Representante LIKE '%" & Busqueda & "%' "
              Case "Emails"
                   sSQL = sSQL _
                        & "WHERE (Email+Email2+EmailR) LIKE '%" & Busqueda & "%' "
              Case Else
                   sSQL = sSQL _
                        & "WHERE " & Campo & " LIKE '%" & Busqueda & "%' "
            End Select
         Else
            sSQL = sSQL & "WHERE Codigo <> '.' "
         End If
         Busqueda = ""
  End Select
  'MsgBox PorExcel & vbCrLf & vbCrLf & sSQL
  If Modulo = "FACTURACION" And Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL & "ORDER BY TD, Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
  
  If Len(Campo) > 1 Then
     If AdoCliente.Recordset.RecordCount > 0 Then
        FrmPatronBusqueda.Visible = False
        DCCliente.Text = AdoCliente.Recordset.fields("Cliente")
     Else
        MsgBox "No existe Datos que buscar"
        FrmPatronBusqueda.Visible = False
     End If
  Else
    Label2.Caption = " NOMBRE DEL CLIENTE" & Space(25) & "Total Clientes: " & Format$(AdoCliente.Recordset.RecordCount, "000000")
  End If
  'MsgBox sSQL & ".........."
  DCCliente.SetFocus
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
Dim TextoBusqueda1 As String
  DatosNuevos
  TextoBusqueda1 = Replace(TextoBusqueda, "&", "'&'")
  'MsgBox TextoBusqueda1
  LstProductos.Clear
  TxtCodigo = "Ninguno"
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente = '" & TextoBusqueda1 & "' "
  Select_Adodc AdoListCtas, sSQL
  With AdoListCtas.Recordset
   If .RecordCount > 0 Then
    'MsgBox TextoBusqueda1
     If .fields("T") = "N" Then LstProductos.AddItem "Activo" Else LstProductos.AddItem "Inactivo"
     MBFecha = .fields("Fecha")
     MBFechaN = .fields("Fecha_N")
     TxtCodigo = .fields("Codigo")
     TxtApellidosS = .fields("Cliente")
     TxtContacto = .fields("Contacto")
     DCCliente = .fields("Cliente")
     Temp_Cliente = DCCliente
     TxtCI_RUC = .fields("CI_RUC")
     Temp_CI_RUC = TxtCI_RUC
     TxtProfesion = .fields("Profesion")
     TxtActividad = .fields("Actividad")
     TxtRazonSocial = .fields("Representante")
     TxtCasilla = .fields("Casilla")
     TxtEmail = .fields("Email")
     TxtEmail2 = .fields("Email2")
     TxtTelefonoS = .fields("Telefono")
     TxtFAX = .fields("FAX")
     TxtCelular = .fields("Celular")
     TxtDirS = .fields("Direccion")
     TxtNumero = .fields("DirNumero")
     CCiudadS = .fields("Ciudad")
     TxtPlan = .fields("Plan_Afiliado")
     Temp_Plan = TxtPlan
     TxtLugarTrabS = .fields("Lugar_Trabajo")
     TxtDirT = .fields("DireccionT")
     TxtTelefonoT = .fields("TelefonoT")
     TxtNo_Dep = .fields("No_Dep")
     TxtGrupo = .fields("Grupo")
     TxtCalificacion = .fields("Calificacion")
     TxtComision = .fields("Porc_C")
     TxtCredito = .fields("Credito")
     CParteR = .fields("Parte_Relacionada")
     MBCtaxCob = FormatoCodigoCta(.fields("Cta_CxP"))
     CodigoEjecutivo = .fields("Cod_Ejec")
     'TxtDescuento = Format(.Fields("Valor_Descuento"), "#,##0.00")
     If .fields("Especial") Then CheqContEsp.value = 1 Else CheqContEsp.value = 0
     If .fields("RISE") Then CheqRISE.value = 1 Else CheqRISE.value = 0
     If .fields("FA") Then LstProductos.AddItem "Cliente de Facturación"
     If .fields("Asignar_Dr") Then CheqDr.value = 1 Else CheqDr.value = 0
     TipoBenef = .fields("TD")
     Temp_TD = .fields("TD")
     For I = 0 To CEstado.ListCount - 1
      If .fields("Est_Civil") = MidStrg(CEstado.List(I), 1, 1) Then
          CEstado = CEstado.List(I)
      End If
     Next I
     If .fields("Sexo") = "M" Then OpcM.value = True Else OpcF.value = True
     Label6.Caption = "* C.I./R.U.C.  [" & TipoBenef & "]"
     DN = Datos_Nacion(.fields("Ciudad"), "C", .fields("Pais"), .fields("Prov"))
     CProvincia.Text = DN.Provincia
     CNacion = DN.Pais
     If CNacion = "" Then CNacion.Text = "ECUADOR"
     TxtApellidosS.Enabled = False
     TxtRazonSocial.Enabled = False
     Select Case .fields("Tipo_Pasaporte")
       Case "00": CTipoProv.Text = "OTRO"
       Case "01": CTipoProv.Text = "PERSONA NATURAL"
       Case Else: CTipoProv.Text = "SOCIEDAD"
     End Select
     
     Listar_Medidores TxtCodigo
     
     TxtLDirs.Text = ""
     TxtDescuento = "0.00"
     sSQL = "SELECT Tipo_Dato, Fecha_Registro, Direccion, Telefono, Ciudad, Descuento, Cuenta_No " _
          & "FROM Clientes_Datos_Extras " _
          & "WHERE Codigo = '" & TxtCodigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Tipo_Dato IN ('DIRECCION','LIBRETAS') " _
          & "ORDER BY Tipo_Dato, Fecha_Registro DESC, Cuenta_No "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        TxtDescuento = Format(AdoAux.Recordset.fields("Descuento"), "#,##0.00")
        Do While Not AdoAux.Recordset.EOF
           If AdoAux.Recordset.fields("Tipo_Dato") = "DIRECCION" Then
              TxtLDirs.Text = TxtLDirs.Text _
                            & AdoAux.Recordset.fields("Fecha_Registro") & ": " & AdoAux.Recordset.fields("Ciudad") & ", " _
                            & AdoAux.Recordset.fields("Direccion") & ". " & AdoAux.Recordset.fields("Telefono") & vbCrLf
           Else
              'MBCuenta.Text = AdoCuentas.Recordset.Fields("Cuenta_No")
               LstProductos.AddItem "Cta. Ahorro No. " & AdoAux.Recordset.fields("Cuenta_No")
           End If
           AdoAux.Recordset.MoveNext
        Loop
     End If
     
     sSQL = "SELECT TC, Cta " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Codigo = '" & TxtCodigo & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY TC,Cta "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           LstProductos.AddItem "Cta. Contable (" & AdoAux.Recordset.fields("TC") & "): " & AdoAux.Recordset.fields("Cta")
           AdoAux.Recordset.MoveNext
        Loop
     End If
     sSQL = "SELECT Ejecutivo " _
          & "FROM Catalogo_Rol_Pagos " _
          & "WHERE Codigo = '" & TxtCodigo & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then LstProductos.AddItem "Asignado a Rol de Pago"
     
     sSQL = "SELECT TP, Fecha, Credito_No " _
          & "FROM Prestamos " _
          & "WHERE Cuenta_No = '" & TxtCodigo & "' " _
          & "AND TP = 'SUSC' " _
          & "ORDER BY Fecha,Credito_No "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           LstProductos.AddItem "Suscripción: [" & AdoAux.Recordset.fields("Fecha") & "] " & AdoAux.Recordset.fields("Credito_No")
           AdoAux.Recordset.MoveNext
        Loop
     End If
     sSQL = "SELECT TP, Fecha, Credito_No " _
          & "FROM Prestamos " _
          & "WHERE Cuenta_No = '" & TxtCodigo & "' " _
          & "AND TP <> 'SUSC' " _
          & "ORDER BY Fecha,Credito_No "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           LstProductos.AddItem AdoAux.Recordset.fields("TP") & " [" & AdoAux.Recordset.fields("Fecha") & "] " & AdoAux.Recordset.fields("Credito_No")
           AdoAux.Recordset.MoveNext
        Loop
     End If
    'Listamos Ejecutivo
     DCEjec.Text = Ninguno
     If AdoEjec.Recordset.RecordCount > 0 Then
        AdoEjec.Recordset.MoveFirst
        AdoEjec.Recordset.Find ("Codigo = '" & CodigoEjecutivo & "' ")
        If Not AdoEjec.Recordset.EOF Then
           DCEjec.Text = AdoEjec.Recordset.fields("Ejecutivo")
           Temp_Ejecutivo = DCEjec.Text
        End If
     End If
     
     Mifecha = PrimerDiaMes(FechaSistema)
     Dia = Day(Mifecha)
     Mes = Month(Mifecha)
     Anio = Year(Mifecha)
     FechaIni = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
     FechaFin = FechaSistema
     Total = 0: Saldo = 0: Contador = 1
   Else
       MsgBox "No Existe"
   End If
  End With
End Sub

Private Sub CCiudadS_GotFocus()
  MarcarTexto CCiudadS
End Sub

Private Sub CCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CCiudadS_LostFocus()
  DN = Datos_Nacion(CCiudadS, "C", DN.CPais, DN.CProvincia)
'''  MsgBox DN.CCiudad & vbCrLf _
'''       & DN.Codigo & vbCrLf _
'''       & DN.CPais & vbCrLf _
'''       & DN.CProvincia & vbCrLf _
'''       & DN.CRegion & vbCrLf _
'''       & DN.Descripcion & vbCrLf _
'''       & DN.Tipo_Rubro & vbCrLf _
'''       & DN.Pais & vbCrLf _
'''       & DN.Provincia & vbCrLf
End Sub

Private Sub CEstado_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CMedidor_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyInsert Then
     Cuenta_No = UCaseStrg(InputBox("Ingrese el Numero de Medidor: ", "INSERCION DE MEDIDOR", "000000"))
     If IsNumeric(Cuenta_No) Then
        Cuenta_No = Format$(Val(Cuenta_No), "000000")
        sSQL = "DELETE * " _
             & "FROM Clientes_Datos_Extras " _
             & "WHERE Codigo = '" & TxtCodigo & "' " _
             & "AND Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Tipo_Dato = 'MEDIDOR' "
        Ejecutar_SQL_SP sSQL
        SetAdoAddNew "Clientes_Datos_Extras"
        SetAdoFields "T", Normal
        SetAdoFields "Codigo", TxtCodigo
        SetAdoFields "Tipo_Dato", "MEDIDOR"
        SetAdoFields "Cuenta_No", Cuenta_No
        SetAdoUpdate
     End If
     Listar_Medidores TxtCodigo
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     Cuenta_No = CMedidor.Text
     Cuenta_No = Replace(Cuenta_No, vbCrLf, "")
     Mensajes = "Esta seguro de desea Eliminar" & vbCrLf _
              & "El Medidor No. " & Cuenta_No & vbCrLf _
              & "De " & TxtApellidosS
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * " _
             & "FROM Clientes_Datos_Extras " _
             & "WHERE Codigo = '" & TxtCodigo & "' " _
             & "AND Cuenta_No = '" & Cuenta_No & "' " _
             & "AND Tipo_Dato = 'MEDIDOR' "
        'MsgBox sSQL
        Ejecutar_SQL_SP sSQL
     End If
     Listar_Medidores TxtCodigo
  End If
End Sub

Private Sub CNacion_GotFocus()
  MarcarTexto CNacion
  NombreCiudad = CCiudadS.Text
  NombreProvincia = CProvincia.Text
End Sub

Private Sub CNacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CNacion_LostFocus()
  DN = Datos_Nacion(CNacion, "N")
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND CPais = '" & DN.CPais & "' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoAux.Recordset.fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CProvincia.AddItem AdoAux.Recordset.fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CProvincia.AddItem "OTRO"
     CProvincia.Text = "OTRO"
  End If
  If CProvincia.Text <> NombreProvincia Then CProvincia.Text = NombreProvincia
End Sub

Private Sub Command1_Click()
  Unload FClientes
End Sub

Private Sub Command2_Click()
  GrabarCliente
End Sub

Private Sub Command3_Click()
  ListarClientes
End Sub

Private Sub Command4_Click()
  ListarClientes
End Sub

Private Sub Command5_Click()
Dim ActualizarRUC As Boolean
  ActualizarRUC = False
  LblSRI.Caption = ""
  Mensajes = ""
 'Consultamos al SRI el RUC
  If Len(TxtCI_RUC) = 13 Then
     TipoSRI = consulta_RUC_SRI(TxtCI_RUC)
       With TipoSRI
         If Len(.RUC_SRI) > 1 Then Mensajes = Mensajes & "R.U.C.: " & .RUC_SRI & vbCrLf
         If Len(.RazonSocial) > 1 Then Mensajes = Mensajes & "RAZON SOCIAL: " & .RazonSocial & vbCrLf
         If Len(.NombreComercial) > 1 Then Mensajes = Mensajes & "NOMBRE COMERCIAL: " & .NombreComercial & vbCrLf
         If Len(.TipoRUC) > 1 Then Mensajes = Mensajes & UCaseStrg(.TipoRUC) & ", "
         If Len(.Obligado) > 1 Then Mensajes = Mensajes & .Obligado & " OBLIGADO A LLEVAR CONTABILIDAD" & vbCrLf
         If Len(.ActividadEconomica) > 1 Then Mensajes = Mensajes & "ACTIVIDAD ECONOMICA: " & .ActividadEconomica & vbCrLf
         If Len(.FechaInicio) > 1 Then Mensajes = Mensajes & "INICIO SU ACTIVIDAD EL " & .FechaInicio & vbCrLf
         If Len(.FechaActualización) > 1 Then Mensajes = Mensajes & "R.U.C. ACTUALIZADO EL " & .FechaActualización & vbCrLf
         If Len(.FechaReinicio) > 1 Then Mensajes = Mensajes & "REINICIO DE ACTIVIDADES: " & .FechaReinicio & vbCrLf
         If Len(.Categoria) > 1 And Len(.ClaseRUC) > 1 Then Mensajes = Mensajes & "CATEGORIA: " & .Categoria & ", CLASE: " & .ClaseRUC & vbCrLf
         If Len(.FechaCese) > 1 Then Mensajes = Mensajes & "CESE DE ACTIVIDADES: " & .FechaCese & vbCrLf
         If Len(.MicroEmpresa) > 1 Then Mensajes = Mensajes & "TIPO DE CONTRIBUYENTE: """ & UCaseStrg(.MicroEmpresa) & """ " & vbCrLf
         If Len(.AgenteRetencion) > 1 Then Mensajes = Mensajes & "AGENTE DE RETENCION: """ & UCaseStrg(.AgenteRetencion) & """ " & vbCrLf
         If Len(.Estado) > 1 Then Mensajes = Mensajes & "ESTADO DEL CONTRIBUYENTE: """ & UCaseStrg(.Estado) & """ "
       End With
       LblSRI.Caption = Mensajes
    
    'MsgBox Temp_CI_RUC & " - " & TxtCI_RUC & " ..."
     If Temp_CI_RUC <> TxtCI_RUC Then
        If Cadena <> "" Then
           If Cadena <> TipoSRI.RazonSocial Then ActualizarRUC = True
        Else
           If Len(TipoSRI.RUC_SRI) = 13 And TxtApellidosS <> TipoSRI.RazonSocial Then ActualizarRUC = True
           Cadena = TxtApellidosS
        End If
     Else
        If TxtApellidosS <> TipoSRI.RazonSocial Then ActualizarRUC = True
        Cadena = TxtApellidosS
     End If
     If Tipo_RUC_CI.Tipo_Beneficiario <> "R" Then ActualizarRUC = False
     TipoBenef = Tipo_RUC_CI.Tipo_Beneficiario
     If UCaseStrg(TipoSRI.Estado) <> "ACTIVO" And Len(TipoSRI.Estado) > 1 Then ActualizarRUC = True
     
     If ActualizarRUC Then
        Titulo = "C O N S U L T A   A L   S.R.I. "
        If Not IsNumeric(TipoSRI.RUC_SRI) And TipoSRI.Estado = Ninguno Then
           Mensajes = "ESTE RUC NO ES VALIDO, " _
                    & "VUELVA A INGRESAR UN RUC CORRECTO"
           If Tipo_RUC_CI.Tipo_Beneficiario <> "P" Then
              MsgBox Mensajes, , Titulo
              TxtCI_RUC.SetFocus
           End If
        Else
            TipoSRI.RazonSocial = Replace(TipoSRI.RazonSocial, "&", "Y")
            Mensajes = "ESTE RUC ESTA ASIGNADO A: " & vbCrLf & vbCrLf _
                     & Cadena & vbCrLf & vbCrLf _
                     & "INFORMACION CORRECTA DEL R.U.C. ES: " & vbCrLf _
                     & String(38, "=") & vbCrLf _
                     & TipoSRI.RazonSocial & vbCrLf & vbCrLf _
                     & "DESEA REGISTRARLE AL CONTRIBUYENTE?"
            If BoxMensaje = vbYes Then
               TxtApellidosS = TipoSRI.RazonSocial
               Command2.SetFocus
            Else
               TxtTelefonoS.SetFocus
            End If
        End If
   
     End If
   Else
     LblSRI.Caption = vbCrLf & vbCrLf & vbCrLf & "          EL DATO INGRESADO ES UNA CEDULA O PASAPORTE" & vbCrLf & vbCrLf & vbCrLf _
                  & "          POR LO TANTO NO SE PRESENTARA DATOS DEL CONTRIBUYENTE"
     TxtTelefonoS.SetFocus
   End If
End Sub

Private Sub Command6_Click()
     RatonNormal
     FrmPatronBusqueda.Visible = True
     LstCampos.SetFocus
End Sub

Private Sub CParteR_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CProvincia_GotFocus()
  MarcarTexto CProvincia
End Sub

Private Sub CProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CProvincia_LostFocus()
  DN = Datos_Nacion(CProvincia, "P", DN.CPais)
  CCiudadS.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & DN.CPais & "' " _
       & "AND CProvincia = '" & DN.CProvincia & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoAux.Recordset.fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CCiudadS.AddItem AdoAux.Recordset.fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
  If CCiudadS.Text <> NombreCiudad Then CCiudadS.Text = NombreCiudad
End Sub

Private Sub CTipoPersona_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CTipoPersona_LostFocus()
  ListarClientes
End Sub

Private Sub CTipoProv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyF10 Then
     With AdoListCtas.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & DCCliente.Text & "' ")
          If Not .EOF Then
             Cadena = .fields("Cliente") & vbCrLf & vbCrLf _
                    & "Codigo: " & .fields("Codigo") & vbCrLf & vbCrLf _
                    & "NUEVO VALOR:"
             Codigo = InputBox(Cadena, "CAMBIO DE CODIGO:", .fields("Codigo"))
             If Codigo <> "" Then
               .fields("Codigo") = UCaseStrg(MidStrg(Codigo, 1, 10))
               .Update
                MsgBox "Codigo Cambiado Con Exito"
             End If
          End If
      End If
     End With
  End If
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 CI_RUC, Cliente " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '" & Busqueda & "%' "
       If Modulo = "FACTURACION" Then
          If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
       ElseIf Modulo = "ROL PAGOS" Then
          sSQL = sSQL & "AND TD IN ('C','P') "
       Else
          sSQL = sSQL & "AND Codigo <> '.' "
       End If
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoCliente, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
  ListarCuenta DCCliente
  TipoDoc = "M"
  LblSRI.Caption = ""
End Sub

Private Sub DCEjec_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

''Private Sub DGClientes_KeyDown(KeyCode As Integer, Shift As Integer)
''  Keys_Especiales Shift
''  Select Case KeyCode
''    Case vbKeyF11
''         DGClientes.Visible = False
''         With AdoListCtas.Recordset
''          If .RecordCount > 0 Then
''              Contador = 0
''             .MoveFirst
''              Do While Not .EOF
''                 Contador = Contador + 1
''                 FClientes.Caption = Contador & " / " & .RecordCount
''                 Cadena = .Fields("CI_RUC")
''                 'MsgBox Cadena
''                 Cadena1 = Cadena
''                 If Len(Cadena) <= 1 Then Cadena = NumEmpresa & Format$(Contador, "00000")
''                 Cadena = CompilarRUC_CI(Cadena)
''                 Si_No = False
''                 If .Fields("CI_RUC") <> TrimStrg(Cadena) Then
''                    .Fields("CI_RUC") = TrimStrg(Cadena)
''                     Si_No = True
''                 End If
''                 ' MsgBox "Desp. " & Cadena
''                 NombreCliente = TrimStrg(UCaseStrg(.Fields("Cliente")))
''                 DireccionCli = TrimStrg(UCaseStrg(.Fields("Direccion")))
''                 If DireccionCli = Ninguno Then DireccionCli = "SD"
''                 If MidStrg(NombreCliente, 1, 1) = "." Then NombreCliente = MidStrg(NombreCliente, 2, Len(NombreCliente))
''                 If MidStrg(NombreCliente, Len(NombreCliente), 1) = "." Then NombreCliente = MidStrg(NombreCliente, 1, Len(NombreCliente) - 1)
''
''                 If MidStrg(DireccionCli, 1, 1) = "." Then DireccionCli = MidStrg(DireccionCli, 2, Len(DireccionCli))
''                 If MidStrg(DireccionCli, Len(DireccionCli), 1) = "." Then DireccionCli = MidStrg(DireccionCli, 1, Len(DireccionCli) - 1)
''
''                .Fields("Cliente") = TrimStrg(NombreCliente)
''                .Fields("Direccion") = TrimStrg(DireccionCli)
''                .Fields("T") = Normal
''                .Fields("Ciudad") = UCaseStrg(.Fields("Ciudad"))
''                 Select Case MidStrg(Cadena, 3, 1)
''                   Case "0" To "6", "8": If Len(Cadena) = 13 Then Cadena = "R" Else Cadena = "C"
''                   Case "7": Cadena = "P"
''                   Case "9": Cadena = "R"
''                   Case Else: Cadena = "O"
''                 End Select
''                 If Len(Cadena1) < 10 Then Cadena = "O"
''                 If Val(MidStrg(Cadena1, 1, 1)) >= 3 Then Cadena = "O"
''                 If .Fields("TD") <> Cadena Then
''                    .Fields("TD") = Cadena
''                     Si_No = True
''                 End If
''                 If Len(.Fields("Telefono")) <= 2 Then .Fields("Telefono") = "000000000"
''                 If Len(.Fields("Celular")) <= 2 Then .Fields("Celular") = "0000000000"
''                 If Len(.Fields("FAX")) <= 2 Then .Fields("FAX") = "000000000"
''                 If Len(.Fields("Direccion")) <= 2 Then .Fields("Direccion") = "SD"
''                 If Len(.Fields("DirNumero")) <= 2 Then .Fields("DirNumero") = "SN"
''                .Update
''                .MoveNext
''              Loop
''          End If
''         End With
''         DGClientes.Visible = True
''    Case vbKeyF12
''  End Select
''  If CtrlDown And KeyCode = vbKeyF1 Then
''     DGClientes.Visible = False
''
''     DGClientes.Visible = True
''  End If
''  If CtrlDown And KeyCode = vbKeyF5 Then DGClientes.AllowUpdate = True
''  If CtrlDown And KeyCode = vbKeyP Then
''     DGClientes.Visible = False
''     If OpcCli.value Then
''        MensajeEncabData = "LISTADO DE CLIENTES FACTURACION"
''     ElseIf OpcCxC.value Then
''            MensajeEncabData = "LISTADO DE CLIENTES POR CONTABILIDAD"
''     ElseIf OpcCxP.value Then
''            MensajeEncabData = "LISTADO DE PROVEEDORES POR CONTABILIDAD"
''     ElseIf OpcLib.value Then
''            MensajeEncabData = "LISTADO DE SOCIOS LIBRETAS"
''     Else
''         MensajeEncabData = "LISTADO GENERAL DE BENEFICIARIOS"
''     End If
''     Imprimir_Clientes AdoListCtas
''     DGClientes.Visible = True
''  End If
''
''End Sub

Private Sub Form_Activate()
  If Modulo = "FACTURACION" Then Actualizar_Datos_Representantes_SP
  
  Actualiza_Buses = Leer_Campo_Empresa("Actualizar_Buses")

  sSQL = "SELECT Codigo, Ejecutivo, Porc_Com " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Ejecutivo "
  SelectDB_Combo DCEjec, AdoEjec, sSQL, "Ejecutivo"
 
  TxtCodigo = "Ninguno"
  
  FormatoMaskCta MBCtaxCob
  
  LstCampos.Clear
  LstCampos.AddItem "Codigo"
  LstCampos.AddItem "Contacto"
  LstCampos.AddItem "CI_RUC_Representante"
  LstCampos.AddItem "Representante"
  LstCampos.AddItem "Emails"
  LstCampos.AddItem "Ciudad"
  LstCampos.AddItem "Grupo"
  LstCampos.AddItem "Prov"
  LstCampos.AddItem "Plan_Afiliado"
  LstCampos.Text = "Codigo"
  
  CTipoProv.Clear
  CTipoProv.AddItem "OTRO"
  CTipoProv.AddItem "PERSONA NATURAL"
  CTipoProv.AddItem "SOCIEDAD"
  CTipoProv.Text = "PERSONA NATURAL"
  
  CParteR.Clear
  CParteR.AddItem "SI"
  CParteR.AddItem "NO"
  CParteR.Text = "NO"
    
  CEstado.Clear
  CEstado.AddItem "Casado"
  CEstado.AddItem "Divorciado"
  CEstado.AddItem "Soltero"
  CEstado.AddItem "Viudo"
  CEstado.AddItem "Otro"
  CEstado.Text = "Soltero"
  
  CTipoPersona.Clear
  CTipoPersona.AddItem "Todos"
  CTipoPersona.AddItem "R.U.C."
  CTipoPersona.AddItem "Cedulas"
  CTipoPersona.AddItem "Pasaporte"
  CTipoPersona.AddItem "Clientes"
  CTipoPersona.AddItem "Clientes Varones"
  CTipoPersona.AddItem "Clientes Mujeres"
  CTipoPersona.AddItem "Clientes Debitos"
  CTipoPersona.AddItem "Clientes Transaferencias"
  CTipoPersona.AddItem "Clientes Descuentos"
  CTipoPersona.AddItem "Clientes Sin Email"
  CTipoPersona.AddItem "Clientes Sin Representantes"
  CTipoPersona.AddItem "Empleados"
  CTipoPersona.AddItem "RISE"
  CTipoPersona.AddItem "Contribuyente Especial"
  CTipoPersona.AddItem "Cuentas por Cobrar"
  CTipoPersona.AddItem "Cuentas por Pagar"
  CTipoPersona.AddItem "Libretas de Ahorro"
  CTipoPersona.Text = "Todos"
  Pagina = 0
  TxtApellidosS = "CONSUMIDOR FINAL"
  FClientes.Caption = "CREACION DEL CLIENTE"
  Label15.Caption = " PLAN AFILIACION"
  If NombreUsuario = "Administrador de Red" Then TBarCliente.buttons("RenumFactMult").Enabled = True
  Select Case Modulo
    Case "CONTABILIDAD"
         TBarCliente.buttons("Imprimir").Enabled = False
         TBarCliente.buttons("Bloquear").Enabled = False
         TBarCliente.buttons("Activar").Enabled = False
         TBarCliente.buttons("Suscripcion").Enabled = False
    Case "ADUANAS"
         TBarCliente.buttons("Imprimir").Enabled = False
         TBarCliente.buttons("Bloquear").Enabled = False
         TBarCliente.buttons("Activar").Enabled = False
         TBarCliente.buttons("CxC").Enabled = False
         TBarCliente.buttons("Ahorros").Enabled = False
         TBarCliente.buttons("RolPago").Enabled = False
         TBarCliente.buttons("Tarjetas").Enabled = False
         Label25.Caption = " FIRMA COMERCIAL"
         Label24.Caption = " AGENTE"
    Case "FACTURACION"
         TBarCliente.buttons("Imprimir").Enabled = False
         TBarCliente.buttons("Bloquear").Enabled = False
         TBarCliente.buttons("Activar").Enabled = False
         TBarCliente.buttons("CxC").Enabled = False
         TBarCliente.buttons("CxP").Enabled = False
         TBarCliente.buttons("Ahorros").Enabled = False
         TBarCliente.buttons("Facturacion").Enabled = False
         'TBarCliente.Buttons("RolPago").Enabled = False
         TBarCliente.buttons("Tarjetas").Enabled = False
         Label15.Caption = " SECTORIZACION"
          
         Label39.Visible = True
         Label29.Visible = True
         DCEjec.Visible = True
         MBCtaxCob.Visible = True
    Case "CAJACREDITO"
         TBarCliente.buttons("CxC").Enabled = False
         TBarCliente.buttons("CxP").Enabled = False
         TBarCliente.buttons("RolPago").Enabled = False
         'TBarCliente.Buttons("FactMult").Enabled = False
         TBarCliente.buttons("Suscripcion").Enabled = False
         TBarCliente.buttons("Facturacion").Enabled = False
  End Select
    
  CNacion.Clear
  sSQL = "SELECT Descripcion_Rubro " _
       & "FROM Tabla_Naciones " _
       & "WHERE TR = 'N' " _
       & "AND Descripcion_Rubro <> 'OTRO' " _
       & "ORDER BY Descripcion_Rubro,CPais "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
         'CNacion.AddItem .Fields("CPais") & " " & .Fields("Descripcion_Rubro")
          CNacion.AddItem .fields("Descripcion_Rubro")
         .MoveNext
       Loop
   End If
  End With
  CNacion.AddItem "OTRO"
  CNacion.Text = "ECUADOR"
  CProvincia.Clear
  sSQL = "SELECT Descripcion_Rubro " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND Descripcion_Rubro <> 'OTRO' " _
       & "ORDER BY Descripcion_Rubro,CProvincia "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CProvincia.AddItem .fields("Descripcion_Rubro")
         .MoveNext
       Loop
   End If
  End With
  CProvincia.AddItem "OTRO"
  CProvincia.Text = "PICHINCHA"
  
  sSQL = "SELECT TOP 50 Codigo, CI_RUC, Cliente " _
       & "FROM Clientes "
  If Modulo = "FACTURACION" Then
     If Mas_Grupos Then sSQL = sSQL & "WHERE DirNumero = '" & NumEmpresa & "' "
  ElseIf Modulo = "ROL PAGOS" Then
     sSQL = sSQL & "WHERE TD IN ('C','P') "
  Else
     sSQL = sSQL & "WHERE Codigo <> '.' "
  End If
  sSQL = sSQL _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
  
  DCCliente.Text = "CONSUMIDOR FINAL"
  'ProgBar.width = MDI_X_Max - 100
  LblSRI.width = MDI_X_Max - LblSRI.Left - 100
  SSTab1.width = MDI_X_Max - 100
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 50
  SSTab1.Tab = 1
  DGCliente.width = SSTab1.width - 200
  DGCliente.Height = SSTab1.Height - DGCliente.Top - 500
  'AdoCodigos.Top = SSTab1.Top + SSTab1.Height - 450
  SSTab1.Tab = 0
  Command1.Top = SSTab1.Top + SSTab1.Height - 450
  Command3.Top = SSTab1.Top + SSTab1.Height - 450
  Command4.Top = SSTab1.Top + SSTab1.Height - 450
  Command2.Top = SSTab1.Top + SSTab1.Height - 450
 
  If Bloquear_Control Then
     TBarCliente.buttons("Imprimir").Enabled = False
     TBarCliente.buttons("Nuevo").Enabled = False
     TBarCliente.buttons("Modificar").Enabled = False
     TBarCliente.buttons("Eliminar").Enabled = False
     TBarCliente.buttons("Bloquear").Enabled = False
     TBarCliente.buttons("Activar").Enabled = False
     TBarCliente.buttons("Grabar").Enabled = False
     TBarCliente.buttons("CxC").Enabled = False
     TBarCliente.buttons("CxP").Enabled = False
     TBarCliente.buttons("Ahorros").Enabled = False
     TBarCliente.buttons("RolPago").Enabled = False
     TBarCliente.buttons("RenumFactMult").Enabled = False
     TBarCliente.buttons("Suscripcion").Enabled = False
     TBarCliente.buttons("Facturacion").Enabled = False
     TBarCliente.buttons("Emails").Enabled = False
     TBarCliente.buttons("UnFacMult").Enabled = False
     TBarCliente.buttons("Archivo_Excel").Enabled = False
     TBarCliente.buttons("Generar_Educativo").Enabled = False
     TBarCliente.buttons("Tarjetas").Enabled = False
     Command2.Enabled = False
  End If
  RatonNormal
  FClientes.WindowState = vbMaximized
  If Nuevo Then
     TxtApellidosS = NombreCliente
     TxtCodigo = "Ninguno"
     TxtGrupo = NumEmpresa
     TxtCI_RUC.SetFocus
  Else
    'ListarCuenta DCCliente.Text
     DCCliente.SetFocus
  End If
End Sub

Private Sub Form_Deactivate()
  FClientes.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoEjec
   ConectarAdodc AdoBuses
   ConectarAdodc AdoCuentas
'   ConectarAdodc AdoCodigos
   ConectarAdodc AdoCliente
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoTarjetas
   ConectarAdodc AdoCreditos
End Sub

Private Sub LstCampos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     FrmPatronBusqueda.Visible = False
     ListarClientes
     DCCliente.SetFocus
  End If
  PresionoEnter KeyCode
End Sub

Private Sub LstCampos_LostFocus()
  TxtCIRUC.SetFocus
End Sub

Private Sub LstProductos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  Contrato_No = SinEspaciosDer(LstProductos.Text)
  If CtrlDown And KeyCode = vbKeyS Then
     Cadena = UCaseStrg(InputBox("Escriba el Nuevo Sector: ", "ACTUALIZACION DE DATOS", ""))
     If Len(Cadena) > 1 Then
        Cadena = MidStrg(Cadena, 1, 8)
        sSQL = "UPDATE Prestamos " _
             & "SET Sector = '" & Cadena & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        Ejecutar_SQL_SP sSQL
     End If
  End If
  If CtrlDown And KeyCode = vbKeyD Then
     Cadena = UCaseStrg(InputBox("Atención a: ", "ACTUALIZACION DE DATOS", ""))
     If Len(Cadena) > 1 Then
        Cadena = MidStrg(Cadena, 1, 40)
        sSQL = "UPDATE Prestamos " _
             & "SET Atencion = '" & Cadena & "' " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Credito_No = '" & Contrato_No & "' "
        Ejecutar_SQL_SP sSQL
     End If
  End If
  If CtrlDown And KeyCode = vbKeyX Then
     msg = UCaseStrg(InputBox("Motivo de la Anulación: ", "ACTUALIZACION DE DATOS", ""))
     Control_Procesos Normal, msg
     sSQL = "UPDATE Clientes_Datos_Extras " _
          & "SET T = '" & Anulado & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Tipo_Dato = 'LIBRETAS' " _
          & "AND Cuenta_No = '" & Contrato_No & "' "
     Ejecutar_SQL_SP sSQL
  End If
  If CtrlDown And KeyCode = vbKeyA Then
     msg = UCaseStrg(InputBox("Motivo de la Activación: ", "ACTUALIZACION DE DATOS", ""))
     Control_Procesos Normal, msg
     sSQL = "UPDATE Clientes_Datos_Extras " _
          & "SET T = '" & Normal & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Tipo_Dato = 'LIBRETAS' " _
          & "AND Cuenta_No = '" & Contrato_No & "' "
     Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Clientes_Facturacion " _
          & "SET T = '" & Normal & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Codigo = '" & TxtCodigo & "' "
     Ejecutar_SQL_SP sSQL
     sSQL = "UPDATE Clientes " _
          & "SET T = '" & Normal & "' " _
          & "WHERE Codigo = '" & TxtCodigo & "' "
     Ejecutar_SQL_SP sSQL
     MsgBox "Proceso exitoso"
  End If
  If CtrlDown And KeyCode = vbKeyL Then
     If SinEspaciosIzq(LstProductos.Text) = "Cta. Ahorro No." Then
        CodigoL = SinEspaciosDer(LstProductos.Text)
        Mensajes = "Imprimir " & TxtApellidosS & "."
        Titulo = "Pregunta de Impresion"
        If BoxMensaje = vbYes Then Imprimir_Apertura CodigoL
     End If
  End If
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub MBFechaN_GotFocus()
  MarcarTexto MBFechaN
End Sub

Private Sub MBFechaN_LostFocus()
  FechaValida MBFechaN
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
Dim Campos_Clientes As String
Dim ConDeudaPendiente As Boolean
'MsgBox Button.key
 Select Case Button.key
   Case "Salir"
        Unload FClientes
   Case "Imprimir"
        Mensajes = "Imprimir " & TxtApellidosS & "."
        Titulo = "Pregunta de Impresion"
        'If BoxMensaje = vbYes Then Imprimir_Apertura MBCuenta
   Case "Nuevo"
        'CTipoPersona.Text = "Todos"
        Command2.Enabled = True
        'ListarClientes
        DatosNuevos
        MsgBox "Ingrese los Datos del Nuevo Beneficiario"
        TxtCI_RUC.SetFocus
   Case "Modificar"
        TxtApellidosS.Enabled = True
        TxtRazonSocial.Enabled = True
        TxtApellidosS.SetFocus
   Case "Eliminar"
        Mensajes = "Esta seguro que desea Eliminar" & vbCrLf _
                 & TxtApellidosS & "."
        Titulo = "Pregunta de grabación"
        If BoxMensaje = vbYes Then
           Si_No = True
           sSQL = "SELECT * " _
                & "FROM Comprobantes " _
                & "WHERE Codigo_B = '" & TxtCodigo & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           Select_Adodc AdoAux, sSQL
           If AdoAux.Recordset.RecordCount > 0 Then Si_No = False
           sSQL = "SELECT * " _
                & "FROM Prestamos " _
                & "WHERE Cuenta_No = '" & TxtCodigo & "' "
           Select_Adodc AdoAux, sSQL
           If AdoAux.Recordset.RecordCount > 0 Then Si_No = False
           sSQL = "SELECT * " _
                & "FROM Trans_Retenciones " _
                & "WHERE Codigo = '" & TxtCodigo & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' "
           'Select_Adodc AdoAux, sSQL
           'If AdoAux.Recordset.RecordCount > 0 Then Si_No = False
           If Si_No Then
              Control_Procesos Normal, "Eliminar a " & TxtApellidosS
              sSQL = "DELETE * " _
                   & "FROM Clientes " _
                   & "WHERE Codigo = '" & TxtCodigo & "' "
              Ejecutar_SQL_SP sSQL
              sSQL = "DELETE * " _
                   & "FROM Clientes_Datos_Extras " _
                   & "WHERE Codigo = '" & TxtCodigo & "' "
              Ejecutar_SQL_SP sSQL
              sSQL = "DELETE * " _
                   & "FROM Clientes_Matriculas " _
                   & "WHERE Codigo = '" & TxtCodigo & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumEmpresa & "' "
              Ejecutar_SQL_SP sSQL
              RatonNormal
              ListarClientes
           Else
              MsgBox "No se puede eliminar este Codigo, existen datos procesados."
           End If
           
        End If
        ListarCuenta TxtApellidosS
   Case "Bloquear"
        Mensajes = "Esta seguro de desea Bloquear" & vbCrLf _
                 & TxtApellidosS & "."
        Titulo = "Pregunta de grabación"
        If BoxMensaje = vbYes Then
           Control_Procesos Normal, "Bloqueo a " & TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET T = 'B' " _
                & "WHERE Codigo = '" & TxtCodigo & "' "
           Ejecutar_SQL_SP sSQL
        End If
        ListarCuenta TxtApellidosS
   Case "Activar"
        Mensajes = "Esta seguro de desea desbloquear" & vbCrLf _
                 & TxtApellidosS & "."
        Titulo = "Pregunta de grabación"
        If BoxMensaje = 6 Then
           Control_Procesos Normal, "Activar a " & TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET T = 'N' " _
                & "WHERE Codigo = '" & TxtCodigo & "' "
           Ejecutar_SQL_SP sSQL
        End If
        ListarCuenta TxtApellidosS
   Case "Grabar"
        GrabarCliente
   Case "CxC"
        If Modulo = "CONTABILIDAD" Then
           Mensajes = "Asignar CxC a " & TxtApellidosS & "."
           Titulo = "Pregunta de CxC"
           If BoxMensaje = vbYes Then
              CodigoCliente = TxtCodigo
              NombreCliente = TxtApellidosS
              SubCta = "C"
              FCxCxP.Show 1
           End If
        Else
           MsgBox "Modulo sin permiso"
        End If
   Case "CxP"
        Mensajes = "Asignar CxP a " & TxtApellidosS & "."
        Titulo = "Pregunta de CxP"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           SubCta = "P"
           FCxCxP.Show 1
        End If
   Case "Ahorros"
        If TxtCodigo = "Ninguno" Then
           MsgBox "No ha grabado el cliente, no se puede asignar Cuenta de Ahorro."
        Else
           Mensajes = "Asignar Cuenta de Ahorros a " & TxtApellidosS & "."
           Titulo = "Pregunta de Creación"
           If BoxMensaje = vbYes Then
              CodigoCli = TxtCodigo
              NombreCliente = TxtApellidosS
              FCtaAhorro.Show 1
           End If
        End If
   Case "RolPago"
        If TipoBenef = "R" Then
           MsgBox "Usted no puede asignar empleados con RUC, solo se permite con Cédula o pasaporte"
        Else
            CodigoPaisEmpleado = DN.CPais
            If TipoBenef <> "C" Then
               MsgBox "Este tipo de Beneficiario" & vbCrLf _
                      & "no es valido en Nomina," & vbCrLf _
                      & "pero se asignará para procesos del Rol"
            End If
            If TxtCodigo = "Ninguno" Then
               MsgBox "No ha grabado el cliente, no se puede asignar a Rol de Pagos"
            Else
               Mensajes = "Asignar Rol de Pagos a " & TxtApellidosS & "."
               Titulo = "Pregunta de Creación"
               If BoxMensaje = vbYes Then
                  FechaValida MBFecha
                  CodigoCli = TxtCodigo
                  NombreCliente = TxtApellidosS
                  NumFacturas = TxtNo_Dep
                  If Modulo = "ROL PAGOS" Then
                     FRolPago.Show 1
                  Else
                     sSQL = "SELECT * " _
                          & "FROM Catalogo_Rol_Pagos " _
                          & "WHERE Codigo = '" & CodigoCli & "' " _
                          & "AND Periodo = '" & Periodo_Contable & "' " _
                          & "AND Item = '" & NumEmpresa & "' "
                     Select_Adodc AdoAux, sSQL
                     If AdoAux.Recordset.RecordCount <= 0 Then SetAddNew AdoAux
                     SetFields AdoAux, "Fecha", MBFecha.Text
                     SetFields AdoAux, "Item", NumEmpresa
                     SetFields AdoAux, "Codigo", CodigoCli
                     SetFields AdoAux, "T", Normal
                     SetFields AdoAux, "SN", "2"
                     SetUpdate AdoAux
                  End If
               End If
            End If
        End If
   Case "RenumFactMult"  ' Renumerar Alumnos de Facturacion
        If ClaveSupervisor Then
           Mensajes = "Este proceso se lo realiza una sola vez al año," & vbCrLf & "desea realizarlo"
           Titulo = "PREGUNTA DE ACTUALIZACION"
           If BoxMensaje = vbYes Then
              sSQL = "SELECT Codigo,CI_RUC,Grupo,Cliente,Direccion,TD " _
                   & "FROM Clientes " _
                   & "WHERE Codigo <> '.' " _
                   & "AND Cliente <> 'CONSUMidOR FINAL' " _
                   & "ORDER BY Cliente "
              Select_Adodc AdoAux, sSQL
              Contador = 0
              RatonReloj
              With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                      Contador = Contador + 1
                      FClientes.Caption = Format$(Contador / .RecordCount, "00%")
                      If .fields("TD") = "O" Then
                         .fields("CI_RUC") = NumEmpresa & Format$(Contador, "00000")
                      Else
                          DigVerif = Digito_Verificador(.fields("CI_RUC"))
                          Caracter = MidStrg(.fields("CI_RUC"), 10, 1)
                          If Tipo_RUC_CI.Tipo_Beneficiario = "O" Then
                            .fields("CI_RUC") = NumEmpresa & Format$(Contador, "000000")
                            .fields("TD") = "O"
                            .Update
                          Else
                             If DigVerif <> Caracter Then
                               .fields("CI_RUC") = NumEmpresa & Format$(Contador, "000000")
                               .fields("TD") = "O"
                               .Update
                             End If
                          End If
                      End If
                     .MoveNext
                   Loop
               End If
              End With
              sSQL = "SELECT Codigo,CI_RUC,Grupo,Cliente,Direccion,TD " _
                   & "FROM Clientes " _
                   & "WHERE FA <> " & Val(adFalse) & " " _
                   & "ORDER BY Grupo,Cliente "
              Select_Adodc AdoAux, sSQL
              Contador = 0: K = 0
              RatonReloj
              With AdoAux.Recordset
               If .RecordCount > 0 Then
                   Do While Not .EOF
                      Contador = Contador + 1
                      FClientes.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
                      DigVerif = Digito_Verificador(.fields("CI_RUC"))
                      Caracter = MidStrg(.fields("CI_RUC"), 10, 1)
                      If Tipo_RUC_CI.Tipo_Beneficiario = "O" Then
                         K = K + 1
                        .fields("CI_RUC") = Format$(K, "00000000")
                        .fields("TD") = "E"
                        .Update
                      Else
                         If DigVerif <> Caracter Then
                            K = K + 1
                           .fields("CI_RUC") = Format$(K, "00000000")
                           .fields("TD") = "E"
                           .Update
                         End If
                      End If
                     .MoveNext
                   Loop
               End If
              End With
              RatonNormal
              MsgBox "Fin del Proceso Puede Generar el Archivo al Banco" & vbCrLf
              FClientes.Caption = "CREACION DEL CLIENTE"
              ListarClientes
           End If
        End If
   Case "Suscripcion"
        Mensajes = "Asignar Contrato de Suscripción a " & TxtApellidosS & "."
        Titulo = "Pregunta de Suscripción"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           Factura_No = 0
           'FSuscripcion.Show 1
        End If
   Case "Facturacion"
        Mensajes = UCaseStrg("Asignar a Facturacion " & TxtApellidosS & "," & vbCrLf & "Asignado con el codigo: " & TxtCodigo)
        Titulo = "Pregunta de Facturacion"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adTrue) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
           MsgBox "PROCESO EXITOSO"
        End If
   Case "Emails"
        ListarEmailsClientes
   Case "UnFacMult"
        Mensajes = "Desactivar de Facturacion a: " & TxtApellidosS & "."
        Titulo = "Pregunta de Facturacion Multiple"
        If BoxMensaje = vbYes Then
           CodigoCliente = TxtCodigo
           NombreCliente = TxtApellidosS
           sSQL = "UPDATE Clientes " _
                & "SET FA = " & Val(adFalse) & " " _
                & "WHERE Codigo = '" & CodigoCliente & "' "
           Ejecutar_SQL_SP sSQL
        End If
   Case "Archivo_Excel"
        ConDeudaPendiente = False
        Mensajes = "Descargar Deuda Pendoente?"
        Titulo = "Pregunta de Facturacion CxC"
        If BoxMensaje = vbYes Then ConDeudaPendiente = True
       'Listado de Codigos para Facturacion de Bancos
        GenerarDataTexto FClientes, AdoCliente
        'ListarClientes LstCampos, TxtCIRUC, True
        Select Case Modulo
          Case "FACTURACION"
               DGCliente.Caption = "LISTADO DE CLIETNES DE FACTURACION ("
          Case Else
               DGCliente.Caption = "LISTADO DE " & CTipoPersona & " LOS BENEFICIARIOS ("
        End Select
        'DGCodigos.Caption = DGCodigos.Caption & AdoCodigos.Recordset.RecordCount & ")"
        'GenerarDataTexto FClientes, AdoCodigos
        If ConDeudaPendiente Then
            sSQL = "SELECT C.T,CI_RUC As COD_BANCO,Cliente As ALUMNOS,Grupo,Direccion As CURSO,C.Fecha_N,DireccionT," _
                 & "Telefono,Sexo,CM.Representante As RAZON_SOCIAL,CM.Cedula_R As RUC_CI,Plan_Afiliado As BUS_No,Archivo_Foto,C.Codigo " _
                 & "FROM Clientes As C,Clientes_ as CM " _
                 & "WHERE CM.Item = '" & NumEmpresa & "' " _
                 & "AND CM.Periodo = '" & Periodo_Contable & "' " _
                 & "AND C.FA <> " & Val(adFalse) & " " _
                 & "AND C.Codigo = CM.Codigo " _
                 & "ORDER BY C.Cliente, C.Grupo "
            Select_Adodc AdoAux, sSQL
            GenerarDataTexto FClientes, AdoAux
        End If
        ListarClientes LstCampos, TxtCIRUC
   Case "Generar_Educativo"
        sSQL = "SELECT C.T,CI_RUC As COD_BANCO,Cliente As ALUMNOS,Grupo,Direccion As CURSO,C.Fecha_N,DireccionT," _
             & "Telefono,Sexo,CM.Representante As RAZON_SOCIAL,CM.Cedula_R As RUC_CI,Plan_Afiliado As BUS_No,Archivo_Foto,C.Codigo " _
             & "FROM Clientes As C,Clientes_Matriculas as CM " _
             & "WHERE CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "AND C.FA <> " & Val(adFalse) & " " _
             & "AND C.Codigo = CM.Codigo " _
             & "ORDER BY C.Cliente, C.Grupo "
        Select_Adodc AdoAux, sSQL
        GenerarDataTexto FClientes, AdoAux
   Case "Cambio_Ejecutivo"
        Mensajes = "Desea cambiar ejecutivos de Venta?"
        Titulo = "CAMBIOS DE EJECUTIVOS DE VENTA"
        If BoxMensaje = vbYes Then
           CodigoEjecutivo = DCEjec
           Cta = CambioCodigoCta(MBCtaxCob)
           FCambioEjecutivo.Show 1
        End If
   Case "Activar_CE"
        sSQL = "UPDATE Clientes " _
             & "SET Clave = TRIM(SUBSTRING(CI_RUC,1,10)) " _
             & "WHERE LEN(Clave) <= 1 "
        Ejecutar_SQL_SP sSQL
        MsgBox "Proceso Terminado"
 End Select

 If Button.key <> "Salir" Then FClientes.Caption = "CREACION DEL CLIENTE"
End Sub

Private Sub TxtActividad_GotFocus()
  MarcarTexto TxtActividad
End Sub

Private Sub TxtActividad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtActividad_LostFocus()
  TextoValido TxtActividad
End Sub

Private Sub TxtCI_RUC_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Letras_Numeros(KeyAscii)
End Sub

Private Sub TxtCIRUC_GotFocus()
  MarcarTexto TxtCIRUC
End Sub

Private Sub TxtCIRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     FrmPatronBusqueda.Visible = False
     DCCliente.SetFocus
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtCIRUC_LostFocus()
   TextoValido TxtCIRUC, , True
   If TxtCIRUC <> Ninguno Then
      ListarClientes LstCampos, TxtCIRUC
   Else
      MsgBox "No existe Datos que buscar"
      ListarClientes
   End If
End Sub

Private Sub TxtComision_GotFocus()
  MarcarTexto TxtComision
End Sub

Private Sub TxtComision_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtComision_LostFocus()
  TextoValido TxtComision, True
End Sub

Private Sub TxtContacto_GotFocus()
   MarcarTexto TxtContacto
End Sub

Private Sub TxtContacto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtContacto_LostFocus()
  TextoValido TxtContacto, , True
End Sub

Private Sub TxtCredito_GotFocus()
  MarcarTexto TxtCredito
End Sub

Private Sub TxtCredito_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCredito_LostFocus()
  TextoValido TxtCredito, True, , 0
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
  TextoValido TxtApellidosS, , True
  If Temp_Cliente <> TxtApellidosS Then
     If Leer_Campo_Cliente("Cliente", TxtApellidosS) <> "" Then
        MsgBox "Este Beneficiario ya está asignado"
        TxtCI_RUC.SetFocus
     End If
  End If
End Sub

Private Sub TxtCasilla_GotFocus()
  MarcarTexto TxtCasilla
End Sub

Private Sub TxtCelular_GotFocus()
  MarcarTexto TxtCelular
End Sub

Private Sub TxtCelular_LostFocus()
  TextoValido TxtCelular, , True
  TxtCelular.Text = Format$(Val(TxtCelular.Text), "0000000000")
End Sub

Private Sub TxtCI_RUC_GotFocus()
  MarcarTexto TxtCI_RUC
  LblSRI = ""
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If AltDown And KeyCode = vbKeyF2 Then TxtCI_RUC.Text = Leer_Codigo_Automatico
  If CtrlDown And KeyCode = vbKeyV Then TxtCI_RUC.Text = Clipboard.GetText()
End Sub

Private Sub TxtCI_RUC_LostFocus()
    TextoValido TxtCI_RUC, , True
    If Len(TxtCI_RUC) > 1 Then
       Mensajes = ""
       DigVerif = Digito_Verificador(TxtCI_RUC)
       If Tipo_RUC_CI.Tipo_Beneficiario = "P" Then
          Mensajes = "Este código es un Pasaporte"
          Titulo = "CONFIRMACION DE PASAPORTE"
          If BoxMensaje <> vbYes Then Tipo_RUC_CI.Tipo_Beneficiario = "O" Else DigVerif = Ninguno
       End If
       If DigVerif = "-" Then
          MsgBox "CEDULA O RUC ERRONEOS, VUELVA A INGRESAR"
          TxtCI_RUC.SetFocus
       Else
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "R": Label4.Caption = "* R A Z O N    S O C I A L"
                      CTipoProv.Text = "OTRO"
                      CParteR.Text = "NO"
                      If Len(TxtApellidosS) <= 1 Then
                         TipoSRI = consulta_RUC_SRI(TxtCI_RUC)
                         TxtApellidosS = TipoSRI.RazonSocial
                      End If
            Case "C": Label4.Caption = "* APELLIDOS Y NOMBRES"
                      CTipoProv.Text = "OTRO"
                      CParteR.Text = "NO"
            Case Else: Label4.Caption = "* BENEFICIARIO"
          End Select
          Label6.Caption = "* C.I./R.U.C.  [" & Tipo_RUC_CI.Tipo_Beneficiario & "]"
       End If
       Cadena = Leer_Campo_Cliente("CI_RUC", TxtCI_RUC)
       If Len(Cadena) > 1 And TxtApellidosS <> Cadena Then
          MsgBox "Este codigo ya esta asignado a:" & vbCrLf & vbCrLf & Cadena & vbCrLf & vbCrLf & "vuelva a ingresar el dato."
          'TxtCI_RUC.SetFocus
       End If
       
    Else
       MsgBox "Este campo no puede tener datos nulos"
       TxtCI_RUC.SetFocus
    End If
End Sub

Private Sub TxtDescuento_GotFocus()
  MarcarTexto TxtDescuento
End Sub

Private Sub TxtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento_LostFocus()
  TextoValido TxtDescuento, True
End Sub

Private Sub TxtDirS_GotFocus()
  MarcarTexto TxtDirS
End Sub

Private Sub TxtDirT_GotFocus()
  MarcarTexto TxtDirT
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
  KeyAscii = Solo_Letras_Numeros(KeyAscii)
End Sub

Private Sub TxtEmail2_KeyPress(KeyAscii As Integer)
  KeyAscii = Solo_Letras_Numeros(KeyAscii)
End Sub

Private Sub TxtFAX_GotFocus()
  MarcarTexto TxtFAX
End Sub

Private Sub TxtGrupo_GotFocus()
  MarcarTexto TxtGrupo
End Sub

Private Sub TxtGrupo_LostFocus()
  TextoValido TxtGrupo, , True
 'If TxtGrupo.Text = Ninguno Then TxtGrupo.Text = NumEmpresa
End Sub

Private Sub TxtLugarTrabS_GotFocus()
  MarcarTexto TxtLugarTrabS
End Sub

Private Sub TxtNo_Dep_GotFocus()
  MarcarTexto TxtNo_Dep
End Sub

Private Sub TxtNumero_GotFocus()
  MarcarTexto TxtNumero
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNumero_LostFocus()
  TextoValido TxtNumero, , True
  If TxtNumero.Text = Ninguno Then TxtNumero.Text = "SN"
End Sub

Private Sub TxtPlan_GotFocus()
  MarcarTexto TxtPlan
End Sub

Private Sub TxtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPlan_LostFocus()
''  If IsDate(FechaIni) And IsDate(FechaFin) And Len(TxtPlan) <= 1 Then
''        FechaIni = Dato_DBF.FechaI
''        FechaFin = Dato_DBF.FechaF
''        Mensajes = "Desea eliminar tambien su historia"
''        Titulo = "Pregunta de Eliminacion"
''        If BoxMensaje = vbYes Then
''           FechaIni = BuscarFecha(FechaIni)
''           FechaFin = BuscarFecha(FechaFin)
''           sSQL = "DELETE * " _
''                & "FROM Clientes_Facturacion " _
''                & "WHERE Codigo = '" & TxtCodigo & "' " _
''                & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
''                & "AND Item = '" & NumEmpresa & "' "
''           Ejecutar_SQL_SP sSQL
''           MsgBox "Proceso exitoso"
''        End If
''  End If
End Sub

Private Sub TxtProfesion_GotFocus()
  MarcarTexto TxtProfesion
End Sub

Private Sub TxtRazonSocial_GotFocus()
  MarcarTexto TxtRazonSocial
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
End Sub

Private Sub TxtEmail2_GotFocus()
  MarcarTexto TxtEmail2
End Sub

Private Sub TxtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail2_LostFocus()
   TextoValido TxtEmail2
   TxtEmail2 = LCase(TxtEmail2)
End Sub

Private Sub TxtCasilla_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCasilla_LostFocus()
  TextoValido TxtCasilla, , True
End Sub

Private Sub TxtDirS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDirS_LostFocus()
  TextoValido TxtDirS, , True
  If TxtDirS.Text = Ninguno Then TxtDirS.Text = "SD"
End Sub

Private Sub TxtDirT_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDirT_LostFocus()
  TextoValido TxtDirT, , True
End Sub

Private Sub TxtFAX_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtFAX_LostFocus()
  TextoValido TxtFAX, , True
  TxtFAX.Text = Format$(Val(TxtFAX.Text), "000000000")
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

Private Sub TxtTelefonoS_GotFocus()
  MarcarTexto TxtTelefonoS
End Sub

Private Sub TxtTelefonoS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoS_LostFocus()
  TextoValido TxtTelefonoS, , True
 'MsgBox MidStrg(TxtTelefonoS, 1, 2)
  If MidStrg(TxtTelefonoS, 1, 2) = "09" Then
     TxtTelefonoS.Text = Format$(Val(TxtTelefonoS.Text), "0000000000")
  Else
     TxtTelefonoS.Text = Format$(Val(TxtTelefonoS.Text), "000000000")
  End If
End Sub

Private Sub TxtTelefonoT_GotFocus()
  MarcarTexto TxtTelefonoT
End Sub

Private Sub TxtTelefonoT_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefonoT_LostFocus()
  TextoValido TxtTelefonoT, , True
  TxtTelefonoT.Text = Format$(Val(TxtTelefonoT.Text), "000000000")
End Sub

Public Sub Listar_Medidores(CodigoDelCliente As String)
 CMedidor.Clear
 If (CodigoDelCliente <> Ninguno) And Len(CodigoDelCliente) > 1 Then
    sSQL = "SELECT Fecha_Registro, Cuenta_No " _
         & "FROM Clientes_Datos_Extras " _
         & "WHERE Codigo = '" & CodigoDelCliente & "' " _
         & "AND Tipo_Dato = 'MEDIDOR' " _
         & "AND T = 'N' " _
         & "ORDER BY Fecha_Registro "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            CMedidor.AddItem .fields("Cuenta_No") & vbCrLf
           .MoveNext
         Loop
     Else
         CMedidor.AddItem "NINGUNO"
     End If
    End With
 Else
    CMedidor.AddItem "NINGUNO"
 End If
 CMedidor.Text = CMedidor.List(0)
End Sub
