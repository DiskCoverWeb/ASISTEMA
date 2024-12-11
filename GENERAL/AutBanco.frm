VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form FRecaudacionBancosCxC 
   BackColor       =   &H00FFFFC0&
   Caption         =   "BANCO BOLIVARIANO"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   Icon            =   "AutBanco.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
   Begin VB.ListBox LstStatud 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   5985
      TabIndex        =   21
      Top             =   8190
      Visible         =   0   'False
      Width           =   5685
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   1588
      ButtonWidth     =   1482
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Boton_Menu"
                  Text            =   "Boton Menu"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Visualizar"
            Key             =   "Visualizar"
            Object.ToolTipText     =   "Visualizar Archivo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Enviar"
            Key             =   "Enviar"
            Object.ToolTipText     =   "Envia rubros de cobros al Banco"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Recibir"
            Key             =   "Recibir"
            Object.ToolTipText     =   "Recibir Abonos del Banco"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSDataListLib.DataCombo DCGrupoF 
         Bindings        =   "AutBanco.frx":0442
         DataSource      =   "AdoGrupo"
         Height          =   360
         Left            =   11340
         TabIndex        =   5
         Top             =   315
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCGrupoI 
         Bindings        =   "AutBanco.frx":0459
         DataSource      =   "AdoGrupo"
         Height          =   360
         Left            =   9450
         TabIndex        =   4
         Top             =   315
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   " CUENTA A LA QUE SE VA ACREDITAR LOS ABONOS"
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
         Left            =   14700
         TabIndex        =   8
         Top             =   105
         Width           =   6105
         Begin MSDataListLib.DataCombo DCBanco 
            Bindings        =   "AutBanco.frx":0470
            DataSource      =   "AdoBanco"
            Height          =   360
            Left            =   105
            TabIndex        =   9
            Top             =   210
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   635
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   192
            Text            =   "Banco"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ORDEN No."
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
         Left            =   13230
         TabIndex        =   6
         Top             =   105
         Width           =   1380
         Begin VB.TextBox TxtOrden 
            BeginProperty Font 
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
            TabIndex        =   7
            Text            =   "0"
            Top             =   210
            Width           =   1170
         End
      End
      Begin VB.CheckBox CheqRangos 
         Caption         =   "Procesar &Por Rangos Grupos:"
         BeginProperty Font 
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
         TabIndex        =   3
         Top             =   0
         Width           =   3690
      End
      Begin VB.Frame Frame3 
         Caption         =   "ENTIDAD FINANCIERA"
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
         Left            =   3465
         TabIndex        =   1
         Top             =   105
         Width           =   5895
         Begin MSDataListLib.DataCombo DCTablaSRI 
            Bindings        =   "AutBanco.frx":0487
            DataSource      =   "AdoTablaSRI"
            Height          =   315
            Left            =   105
            TabIndex        =   2
            Top             =   210
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   4194304
            Text            =   "Banco"
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
      End
   End
   Begin VB.TextBox TxtFile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5580
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   18
      Top             =   2415
      Width           =   10725
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&S"
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
      Left            =   14700
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1050
      Width           =   330
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2730
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutBanco.frx":04A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutBanco.frx":0D7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutBanco.frx":1655
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutBanco.frx":196F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutBanco.frx":2225
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PctBanco 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   105
      ScaleHeight     =   1065
      ScaleWidth      =   9045
      TabIndex        =   19
      Top             =   840
      Width           =   9045
   End
   Begin VB.CheckBox CheqSAT 
      BackColor       =   &H00FF8080&
      Caption         =   "Metodo SAT"
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
      Left            =   11970
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.CheckBox CheqMatricula 
      BackColor       =   &H00FF8080&
      Caption         =   "Generar &Matricula"
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
      Left            =   11970
      TabIndex        =   14
      Top             =   1050
      Width           =   2640
   End
   Begin VB.CheckBox CheqPend 
      BackColor       =   &H00FF8080&
      Caption         =   "Sin &Deuda Pendiente"
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
      Left            =   11970
      TabIndex        =   15
      Top             =   1365
      Width           =   2640
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   105
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   105
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   105
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
      Caption         =   "Producto"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   10605
      TabIndex        =   11
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoAbono 
      Height          =   330
      Left            =   105
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
      Caption         =   "Abono"
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
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   10605
      TabIndex        =   13
      Top             =   1575
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSAdodcLib.Adodc AdoPendiente 
      Height          =   330
      Left            =   105
      Top             =   5775
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
      Caption         =   "Pendiente"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   105
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
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
      Caption         =   "Factura"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   105
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
      Caption         =   "Banco"
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   105
      Top             =   4830
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
      Caption         =   "IngCaja"
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
      Left            =   105
      Top             =   5145
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
   Begin MSAdodcLib.Adodc AdoTablaSRI 
      Height          =   330
      Left            =   105
      Top             =   5460
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
      Caption         =   "TablaSRI"
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
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   4095
      TabIndex        =   22
      Top             =   8190
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgLstFTP"
      SmallIcons      =   "ImgLstFTP"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Archivos"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tamaño"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modificado"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   17
      Top             =   1995
      Width           =   10725
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tope de &Pago"
      BeginProperty Font 
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
      TabIndex        =   12
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Facturacion"
      BeginProperty Font 
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
      TabIndex        =   10
      Top             =   1050
      Width           =   1380
   End
End
Attribute VB_Name = "FRecaudacionBancosCxC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''sSQL = "UPDATE Facturas As F " _
'''     & "SET Saldo_MN = Total_MN - (" _
'''     & "SELECT SUM(Abono) As Abonos " _
'''     & "FROM Trans_Abonos As TA " _
'''     & "WHERE TA.Item = '" & NumEmpresa & "' " _
'''     & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''     & "AND TA.T <> 'A' " _
'''     & "AND F.Item = TA.Item " _
'''     & "AND F.Periodo = TA.Periodo " _
'''     & "AND F.CodigoC = TA.CodigoC " _
'''     & "AND F.Factura = TA.Factura " _
'''     & "AND F.TC = TA.TP " _
'''     & "GROUP BY TA.Factura) "
'''
'''MsgBox "ANTES => " & vbCrLf & sSQL
'''
'''Ejecutar_SQL_SP sSQL
'''MsgBox sSQL

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String

Dim NumFile As Integer
Dim NumFileAbonos As Integer
Dim NumFileDetalle As Integer
Dim NumFileAlumnos As Integer
Dim NumFileFacturas As Integer
Dim NumFileProducto As Integer

Dim RutaGeneraFile As String
Dim RutaGeneraFileAbonos As String
Dim RutaGeneraFileDetalle As String
Dim RutaGeneraFileAlumnos As String
Dim RutaGeneraFileFacturas As String
Dim RutaGeneraFileProducto As String

Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal
Dim RutaBackupXX As String

Dim Costo_Banco As Double
Dim Total_Costo_Banco As Double
Dim Cta_Gasto_Banco As String
Dim Tipo_Carga As Byte

'''Private Sub Generar_Facturas()
'''Dim AuxNumEmp As String
'''Dim DiaV As Integer
'''Dim MesV As Integer
'''Dim AñoV As Integer
'''Dim Total_Alumnos As Long
'''Dim CamposFile() As Campos_Tabla
'''Dim Separador As String
'''Dim Estab As Boolean
'''Dim CaptionTemp As String
'''  CaptionTemp = FRecaudacionBancosCxC.Caption
'''  TextoImprimio = ""
'''  sSQL = "UPDATE Facturas " _
'''       & "SET X = '.' " _
'''       & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Item = '" & NumEmpresa & "' " _
'''       & "AND TC NOT IN ('C','P') " _
'''       & "AND X <> '.' " _
'''       & "AND T <> 'A' "
'''  Ejecutar_SQL_SP sSQL
'''  Separador = ","
'''  FechaValida MBFechaI
'''  Mifecha = BuscarFecha(MBFechaI)
'''  FechaTexto = MBFechaI ' FechaSistema
'''  DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
'''  RutaGeneraFile = UCaseStrg(Dir1.Path & "\" & NombreArchivo)
'''  TotalIngreso = 0
'''  Contador = 0
'''  FileResp = 0
'''  ''ProgBarra.value = 0
'''  'ProgBarra.Min = 0
''' 'Establecemos los campos del archivo plano del Banco
'''  NumFile = FreeFile
'''  Total_Alumnos = 0
'''  Open RutaGeneraFile For Input As #NumFile
'''    Do While Not EOF(NumFile)
'''      'Leemos el encabezado
'''       If Total_Alumnos = 0 Then
'''          Line Input #NumFile, Cod_Field
'''          ReDim CamposFile(5) As Campos_Tabla
'''       End If
'''       Line Input #NumFile, Cod_Field
'''      'Comenzamos la subida de las Facturas y los Abonos
'''       Cadena = Cod_Field
'''       CamposFile(0).Valor = TrimStrg(MidStrg(Cod_Field, 79, 10))      ' Codigo
'''       CamposFile(1).Valor = MidStrg(Cod_Field, 27, 2) & "/" _
'''                           & MidStrg(Cod_Field, 25, 2) & "/" _
'''                           & MidStrg(Cod_Field, 21, 4)             ' Fecha
'''       CamposFile(2).Valor = Val(MidStrg(Cod_Field, 252, 7))       ' Secuencial
'''       CamposFile(3).Valor = Val(MidStrg(Cod_Field, 32, 15)) / 100 ' Valor a pagar
'''      'Actualizamos de que alumnos vamos a ingresar el abono
'''       CodigoCli = Ninguno
'''       Select Case TextoBanco
'''         Case "PACIFICO"
'''              CodigoP = CamposFile(0).Valor
'''              FechaTexto = CamposFile(1).Valor
'''              Factura_No = CamposFile(2).Valor
'''       End Select
'''       Mifecha = BuscarFecha(FechaTexto)
'''      'Actualizamos las Facturas del Alumno
'''       sSQL = "UPDATE Facturas " _
'''            & "SET X = 'D' " _
'''            & "WHERE Factura = " & Factura_No & " " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND Item = '" & NumEmpresa & "' " _
'''            & "AND TC NOT IN ('C','P') "
'''       Ejecutar_SQL_SP sSQL
'''       sSQL = "UPDATE Detalle_Factura " _
'''            & "SET X = 'D' " _
'''            & "WHERE Factura = " & Factura_No & " " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND Item = '" & NumEmpresa & "' "
'''       Ejecutar_SQL_SP sSQL
'''       sSQL = "UPDATE Trans_Abonos " _
'''            & "SET X = 'D' " _
'''            & "WHERE Factura = " & Factura_No & " " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND Item = '" & NumEmpresa & "' "
'''       Ejecutar_SQL_SP sSQL
'''      'Eliminamos los abonos de este dia del Alumno
'''       Total_Alumnos = Total_Alumnos + 1
'''    Loop
'''  Close #NumFile
''' 'Actualizamos los saldos de las facturas
'''  sSQL = "DELETE * " _
'''       & "FROM Facturas " _
'''       & "WHERE X = 'D' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  Ejecutar_SQL_SP sSQL
'''  sSQL = "DELETE * " _
'''       & "FROM Detalle_Factura " _
'''       & "WHERE X = 'D' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  Ejecutar_SQL_SP sSQL
'''  sSQL = "DELETE * " _
'''       & "FROM Trans_Abonos " _
'''       & "WHERE X = 'D' " _
'''       & "AND Periodo = '" & Periodo_Contable & "' " _
'''       & "AND Item = '" & NumEmpresa & "' "
'''  Ejecutar_SQL_SP sSQL
'''  'ProgBarra.Max = Total_Alumnos + 1
''' 'Ingresamos la Factura
'''  TotalIngreso = 0
'''  NumFile = FreeFile
'''  Total_Alumnos = 0
'''  Open RutaGeneraFile For Input As #NumFile
'''    Do While Not EOF(NumFile)
'''      'Leemos el encabezado
'''       If Total_Alumnos = 0 Then
'''          Line Input #NumFile, Cod_Field
'''          ReDim CamposFile(5) As Campos_Tabla
'''          TotalIngreso = 0
'''       End If
'''       Line Input #NumFile, Cod_Field
'''      'Comenzamos la subida de las Facturas y los Abonos
'''       Cadena = Cod_Field
'''       CamposFile(0).Valor = TrimStrg(MidStrg(Cod_Field, 79, 10))      ' Codigo
'''       CamposFile(1).Valor = MidStrg(Cod_Field, 27, 2) & "/" _
'''                           & MidStrg(Cod_Field, 25, 2) & "/" _
'''                           & MidStrg(Cod_Field, 21, 4)             ' Fecha
'''       CamposFile(2).Valor = Val(MidStrg(Cod_Field, 252, 7))       ' Secuencial
'''       CamposFile(3).Valor = Val(MidStrg(Cod_Field, 32, 15)) / 100 ' Valor a pagar
'''      'Actualizamos de que alumnos vamos a ingresar el abono
'''       Select Case TextoBanco
'''         Case "PACIFICO"
'''              CodigoP = CamposFile(0).Valor
'''              FechaTexto = CamposFile(1).Valor
'''              Factura_No = CamposFile(2).Valor
'''              Total = CamposFile(3).Valor
'''       End Select
'''       Mifecha = BuscarFecha(FechaTexto)
'''       CodigoCli = Ninguno
'''       TextoCheque = Ninguno
'''       NombreBanco = "DEPOSITO EFECTIVO"
'''      'MsgBox CodigoP & vbCrLf & FechaTexto & vbCrLf & NombreBanco
'''       Si_No = True
'''       If AdoClientes.Recordset.RecordCount > 0 Then
'''          Do While Len(CodigoP) <= 10 And Si_No
'''             AdoClientes.Recordset.MoveFirst
'''             AdoClientes.Recordset.Find ("CI_RUC = '" & CodigoP & "' ")
'''             If Not AdoClientes.Recordset.EOF Then
'''                CodigoCli = AdoClientes.Recordset.fields("Codigo")
'''                TextoCheque = AdoClientes.Recordset.fields("Grupo")
'''                Si_No = False
'''             Else
'''                CodigoP = "0" & CodigoP
'''             End If
'''          Loop
'''       End If
'''       If Len(CodigoP) > 10 Then CodigoCli = Ninguno
'''      'Insertamos la Factura del Alumno
'''       If Total > 0 Then
'''         'Abonos
'''          SetAdoAddNew "Trans_Abonos"
'''          SetAdoFields "T", Cancelado
'''          SetAdoFields "TP", TipoFactura
'''          SetAdoFields "CodigoC", CodigoCli
'''          SetAdoFields "Fecha", FechaTexto
'''          SetAdoFields "Comprobante", Ninguno
'''          SetAdoFields "Factura", Factura_No
'''          SetAdoFields "Abono", Abono
'''          SetAdoFields "Banco", NombreBanco
'''          SetAdoFields "Cheque", TextoCheque
'''          SetAdoFields "Cta", Cta_Del_Banco     'Cta_CajaG
'''          SetAdoFields "Cta_CxP", Cta_Cobrar
'''          SetAdoUpdate
'''         'Facturas
'''          SetAdoAddNew "Facturas"
'''          SetAdoFields "T", Cancelado
'''          SetAdoFields "TP", TipoFactura
'''          SetAdoFields "CodigoC", CodigoCli
'''          SetAdoFields "Fecha", FechaTexto
'''          SetAdoFields "Factura", Factura_No
'''          SetAdoFields "Total_MN", Total
'''          SetAdoFields "Sin_IVA", Total
'''          SetAdoFields "Subtotal", Total
'''          SetAdoFields "Cta_CxP", Cta_Cobrar
'''          SetAdoUpdate
'''         'Detalle Factura
'''          SetAdoAddNew "Detalle_Factura"
'''          SetAdoFields "T", Cancelado
'''          SetAdoFields "TP", TipoFactura
'''          SetAdoFields "CodigoC", CodigoCli
'''          SetAdoFields "Codigo", "01.01"
'''          SetAdoFields "Fecha", FechaTexto
'''          SetAdoFields "Factura", Factura_No
'''          SetAdoFields "Cantidad", 1
'''          SetAdoFields "Total", Total
'''          SetAdoFields "Precio", Total
'''          SetAdoFields "Producto", "Cobro de Pensiones por el Banco"
'''          SetAdoUpdate
'''          'MsgBox ".........."
'''       End If
'''       TotalIngreso = TotalIngreso + Total
'''       Total_Alumnos = Total_Alumnos + 1
'''       'ProgBarra.value = 'ProgBarra.value + 1
'''    Loop
'''  Close #NumFile
'''  RatonNormal
'''  ''ProgBarra.value = 'ProgBarra.Max
'''  FRecaudacionBancosCxC.Caption = CaptionTemp
'''  MsgBox "ARCHIVO DE ABONO DEL DIA: " & FechaTexto & vbCrLf & vbCrLf _
'''       & "SE ACTUALIZARON: " & Total_Alumnos & " ALUMNOS." & vbCrLf & vbCrLf _
'''       & "EL CIERRE DIARIO DE CAJA ES POR " & Moneda & " " & Format$(TotalIngreso, "#,##0.00") & vbCrLf & vbCrLf _
'''       & "OBTENIDO DEL ARCHIVO: " & vbCrLf & vbCrLf & RutaGeneraFile
'''  'MsgBox "(" & TextoImprimio & ")"
'''  If Len(TextoImprimio) > 2 Then
'''     Unload FRecaudacionBancosCxC
'''     FInfoError.Show
'''  End If
'''End Sub

Private Sub Novedades()
Dim Cont As Integer
Dim CaptionTemp As String
  CaptionTemp = FRecaudacionBancosCxC.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  TipoDoc = ""
  If CheqMatricula.value = 1 Then TipoDoc = "0" Else TipoDoc = "1"
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaFinal = MBFechaF
  FechaTexto = FechaSistema
  FechaTexto1 = Format$(MBFechaI, "MM/dd/yyyy")
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(MBFechaF)
  TextoImprimio = ""
 'Saldo Pendiente de las Facturas
  sSQL = "SELECT TA.CodigoC,C.Grupo,C.Casilla,C.Actividad,C.Cliente,TA.Comprobante,C.CI_RUC,C.Direccion,SUM(Abono) As Abonos " _
       & "FROM Trans_Abonos As TA,Clientes As C " _
       & "WHERE TA.Fecha = #" & Mifecha & "# " _
       & "AND TA.Item = '" & NumEmpresa & "' " _
       & "AND TA.Periodo = '" & Periodo_Contable & "' " _
       & "AND TA.Comprobante IN ('A','E') " _
       & "AND TA.T <> 'A' " _
       & "AND TA.CodigoC = C.Codigo " _
       & "GROUP BY TA.CodigoC,C.Grupo,C.Casilla,C.Actividad,C.Cliente,TA.Comprobante,C.CI_RUC,C.Direccion " _
       & "HAVING SUM(Abono) > 0 " _
       & "ORDER BY TA.CodigoC,C.Actividad,C.Cliente,C.CI_RUC,C.Direccion "
  'MsgBox sSQL
  Select_Adodc AdoFactura, sSQL
  Select Case TextoBanco
    Case "PICHINCHA"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
    Case "BGR_EC"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         
    Case "INTERNACIONAL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         
    Case "BOLIVARIANO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Actualizar_Bolivariano
    Case "PACIFICO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
    Case Else
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         MsgBox "No esta definido este Banco"
  End Select
  FRecaudacionBancosCxC.Caption = CaptionTemp
End Sub

Private Sub Visualizar_Archivo()
'Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog)
'Dir_Dialog.Filename = Guardar_Archivo(Me.hwnd, Dir_Dialog)
'Dir_Dialog.Filter = "Archivo de texto|*.txt|" _
'                  & "Mapa de bits|*.bmp|" _
'                  & "Todos los archivos|*.*"
Dim MaxCar As Integer
Dim Result As String
Label4.Caption = Ninguno

Dir_Dialog.Filter = "Todos los archivos|*.*"
Dir_Dialog.InitDir = RutaSysBases & "\Banco\Abonos\"
Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, OpenFile)
  
  NombreArchivo = Dir_Dialog.File
  RutaGeneraFile = TrimStrg(Dir_Dialog.Filename)
  Label4.Caption = RutaGeneraFile
  If NombreArchivo <> "" Then
     'TxtFile = Leer_Archivo_Texto(RutaGeneraFile)
     RatonReloj
     MaxCar = 0
    'MsgBox RutaGeneraFile
     NumFile = FreeFile: TxtFile = ""
     Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          If Len(Cod_Field) > MaxCar Then MaxCar = Len(Cod_Field)
          TxtFile = TxtFile & Cod_Field & vbCrLf
       Loop
     Close #NumFile
     J = 1: K = 0
     Cadena = ""
     Cadena1 = ""
     For I = 1 To MaxCar
         Cadena = Cadena & CStr(J)
         J = J + 1
         If J > 9 Then
            Cadena = Cadena & "0"
            J = 1
            K = K + 1
            If K <= 10 Then
               Cadena1 = Cadena1 & String(9, " ") & CStr(K)
            Else
               Cadena1 = Cadena1 & String(8, " ") & CStr(K)
            End If
         End If
     Next I
     Cadena = Cadena & vbCrLf
     Cadena1 = Cadena1 & vbCrLf
     
     TxtFile = Cadena1 & Cadena & TxtFile
  Else
     MsgBox "Seleccione un archivo"
  End If
  RatonNormal
End Sub

Private Sub CheqRangos_Click()
  If CheqRangos.value = 0 Then
     DCGrupoI.Visible = False
     DCGrupoF.Visible = False
  Else
     DCGrupoI.Visible = True
     DCGrupoF.Visible = True
  End If
End Sub

Private Sub Command2_Click()
  Unload FRecaudacionBancosCxC
End Sub

Private Sub Enviar_Rubros()
Dim Cont As Integer
Dim CaptionTemp As String
Dim Costo_Banco As Double
Dim Tabulador As String
  
  SumaBancos = 0
  CaptionTemp = FRecaudacionBancosCxC.Caption
  FechaValida MBFechaI
  FechaValida MBFechaF
  TipoDoc = ""
  If CheqMatricula.value = 1 Then TipoDoc = "0" Else TipoDoc = "1"
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaFinal = MBFechaF
  FechaTexto = FechaSistema
  FechaTexto1 = Format$(MBFechaI, "MM/dd/yyyy")
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(MBFechaF)
  TextoImprimio = ""
 'Si LEN(Actividad) es 1 es en efectivo y si es mayor que 3 deposito de cta CTE/AHO.
 'Saldo Pendiente de las Facturas
'''  sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,CI_RUC,C.Direccion,C.Grupo,SUM(Saldo_MN) As Saldo_Pend " _
'''       & "FROM Facturas As F,Clientes As C " _
'''       & "WHERE F.Item = '" & NumEmpresa & "' " _
'''       & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''       & "AND F.T = 'P' " _
'''       & "AND NOT F.TC IN ('C','P') " _
'''       & "AND F.CodigoC = C.Codigo " _
'''       & "GROUP BY C.Grupo,CI_RUC,F.CodigoC,C.Actividad,C.Cliente,C.Direccion " _
'''       & "HAVING SUM(Saldo_MN) > 0 " _
'''       & "ORDER BY C.Grupo,CI_RUC,F.CodigoC,C.Actividad,C.Cliente,C.Direccion "
  
 'Actualizamos las Facturas del Alumno
  
  Eliminar_Nulos_SP "Facturas"
  
  sSQL = "UPDATE Facturas " _
       & "SET Anio_Mes = SUBSTRING(DF.Ticket,1,4) + '-' + RIGHT(REPLICATE('0', 2 - LEN(CAST(DF.Mes_No As VARCHAR))), 2) + CAST(DF.Mes_No As VARCHAR) " _
       & "FROM Facturas As F, Detalle_Factura As DF " _
       & "WHERE F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND LEN(Anio_Mes) = 1 " _
       & "AND LEN(DF.Ticket) = 4 " _
       & "AND F.Periodo = DF.Periodo " _
       & "AND F.Item = DF.Item " _
       & "AND F.TC = DF.TC " _
       & "AND F.Serie = DF.Serie " _
       & "AND F.Factura = DF.Factura "
  Ejecutar_SQL_SP sSQL
       
  sSQL = "SELECT F.CodigoC, C.Actividad, C.Cliente, C.CI_RUC, C.Direccion, C.Grupo, F.Fecha, F.Serie, F.Factura, F.Total_MN, F.Saldo_MN, F.Anio_Mes " _
       & "FROM Facturas As F, Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# " _
       & "AND F.T = 'P' " _
       & "AND F.Saldo_MN > 0 " _
       & "AND NOT F.TC IN ('C','P') "
  If CheqRangos.value <> 0 Then sSQL = sSQL & "AND C.Grupo BETWEEN '" & DCGrupoI & "' and '" & DCGrupoF & "' "
  sSQL = sSQL _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY C.Grupo, C.Cliente, F.Anio_Mes, F.Serie, F.Factura, CI_RUC,F.CodigoC,C.Actividad,C.Direccion, F.Fecha, F.Saldo_MN "
  Select_Adodc AdoPendiente, sSQL
 'Detalle de las Facturas Emitidas del mes
  sSQL = "SELECT DF.*,C.Cliente,C.Grupo,C.CI_RUC,C.Direccion,CP.Item_Banco,CP.Desc_Item " _
       & "FROM Detalle_Factura As DF,Clientes As C,Catalogo_Productos As CP " _
       & "WHERE DF.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# " _
       & "AND DF.Item = '" & NumEmpresa & "' " _
       & "AND DF.Periodo = '" & Periodo_Contable & "' " _
       & "AND DF.T = 'P' " _
       & "AND DF.CodigoC = C.Codigo " _
       & "AND DF.Item = CP.Item " _
       & "AND DF.Periodo = CP.Periodo " _
       & "AND DF.Codigo = CP.Codigo_Inv " _
       & "ORDER BY C.Grupo,C.Cliente, DF.Fecha "
  Select_Adodc AdoDetalle, sSQL
 'Facturas Emitidas del mes
  sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,C.Grupo,CI_RUC,C.Direccion,SUM(Saldo_MN) As Saldo_Pend " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T = 'P' " _
       & "AND NOT F.TC IN ('C','P') " _
       & "AND F.CodigoC = C.Codigo " _
       & "GROUP BY C.Grupo,C.Cliente,CI_RUC,F.CodigoC,C.Actividad,C.Direccion " _
       & "HAVING SUM(Saldo_MN) > 0 " _
       & "ORDER BY C.Grupo,C.Cliente,CI_RUC "
  Select_Adodc AdoFactura, sSQL
 'Facturas Emitidas del mes
  sSQL = "SELECT F.*,C.Cliente,C.Grupo,C.CI_RUC,C.Direccion,C.Casilla " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T <> 'A' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY C.CI_RUC,C.Grupo,F.Fecha "
  Select_Adodc AdoAux, sSQL
  Select Case TextoBanco
    Case "PICHINCHA"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pichincha
    Case "BGR_EC"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_BGR_EC
    Case "INTERNACIONAL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Internacional
    Case "BOLIVARIANO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Bolivariano
    Case "PACIFICO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Pacifico
    Case "PRODUBANCO"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Produbanco
    Case "GUAYAQUIL"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Guayaquil
    Case "COOPJEP"
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         Generar_Coop_Jep
    Case Else
         FechaFin = BuscarFecha(UltimoDiaMes(MBFechaF))
         MsgBox "No esta definido este Banco"
  End Select

   'Facturas Emitidas del mes
   'Generacion del Resumen de la facturacion del mes
    Tabulador = ";"
    RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\RESUMEN_MES_" & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)) & "_" & Cta_Bancaria & ".csv"
    NumFileFacturas = FreeFile
    Contador = 0
    FechaTexto = BuscarFecha(MBFechaI)
   'MsgBox RutaGeneraFile
    TxtFile = ""
    Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
    With AdoPendiente.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Print #NumFileFacturas, "No." & Tabulador;
         Print #NumFileFacturas, "GRUPO" & Tabulador;
         Print #NumFileFacturas, "CODIGO" & Tabulador;
         Print #NumFileFacturas, "BENEFICIARIO" & Tabulador;
         Print #NumFileFacturas, "DETALLE" & Tabulador;
         Print #NumFileFacturas, "AÑO-MES" & Tabulador;
         Print #NumFileFacturas, "SERIE" & Tabulador;
         Print #NumFileFacturas, "FACTURA No" & Tabulador;
         Print #NumFileFacturas, "TOTAL FACTURA" & Tabulador;
         Print #NumFileFacturas, "SALDO FACTURA" & Tabulador
         Do While Not .EOF
            Contador = Contador + 1
            Grupo_No = .fields("Grupo")
            Codigo = .fields("CodigoC")
            CodigoCli = .fields("CI_RUC")
            NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
            Codigo1 = Sin_Signos_Especiales(.fields("Direccion"))
            Codigo2 = .fields("Anio_Mes")
            SerieFactura = "'" & .fields("Serie")
            Factura_No = .fields("Factura")
            Total_Factura = .fields("Total_MN")
            Total_Pagar = .fields("Saldo_MN")
            Total = Total_Factura - Total_Pagar
          ' Empieza la trama por Alumno
            Print #NumFileFacturas, Contador & Tabulador;
            Print #NumFileFacturas, Grupo_No & Tabulador;
            Print #NumFileFacturas, CodigoCli & Tabulador;
            Print #NumFileFacturas, NombreCliente & Tabulador;
            Print #NumFileFacturas, Codigo1 & Tabulador;
            Print #NumFileFacturas, Codigo2 & Tabulador;
            Print #NumFileFacturas, SerieFactura & Tabulador;
            Print #NumFileFacturas, Factura_No & Tabulador;
            Print #NumFileFacturas, Total_Factura & Tabulador;
            Print #NumFileFacturas, Total_Pagar & Tabulador
           .MoveNext
         Loop
     End If
    End With
    Close #NumFileFacturas
    '''ProgBarra.value = 'ProgBarra.Max
    RatonNormal
    FRecaudacionBancosCxC.Caption = CaptionTemp
End Sub

'Recibir Abonos del Banco
Private Sub Recibir_Abonos()
Dim Estab As Boolean
Dim OrdenValida As Boolean

Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer

Dim Total_Alumnos As Long

Dim Total_Dep_Confirmar As Currency
Dim AbonosAnticipados As Currency

Dim CamposFile() As Campos_Tabla

Dim Separador As String
Dim CampoTemp As String
Dim AbonosPar As String
Dim Fecha_Tope As String
Dim Proceso_Ok As String
Dim Orden_Pago As String
Dim AuxNumEmp As String

  FechaValida MBFechaI
  FechaValida MBFechaF
  
  Dir_Dialog.Filter = "Todos los archivos|*.txt"
  Dir_Dialog.InitDir = RutaSysBases & "\Banco\Abonos\"
  If Label4.Caption <> Ninguno Then
     Mensajes = "SUBIR EL ARCHIVO DE LA PANTALLA"
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = vbYes Then
        NombreArchivo = Label4.Caption
     Else
        Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, OpenFile)
        NombreArchivo = Dir_Dialog.File
        RutaGeneraFile = Dir_Dialog.Filename
     End If
  Else
     Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, OpenFile)
     NombreArchivo = Dir_Dialog.File
     RutaGeneraFile = Dir_Dialog.Filename
  End If
  NombreArchivo = Replace(NombreArchivo, vbCrLf, "")
  RutaGeneraFile = Replace(RutaGeneraFile, vbCrLf, "")
  Label4.Caption = Replace(Label4.Caption, vbCrLf, "")
  Progreso_Barra.Mensaje_Box = ""
  Progreso_Iniciar
 'Subo el archivo al Servidor dse DB
  Subir_Archivo_FTP_Linode ftp, LstStatud, LstVwFTP, RutaGeneraFile
 'Proceso el archivo de abonos
  Subir_Archivo_Abonos_Bancos_SP RutaGeneraFile, TextoBanco
'''  MsgBox "proceso terminado"
'''  Exit Sub
  Contador = 0
  CantCampos = 0
  TotalIngreso = 0
  Separador = Ninguno
  Orden_Pago = Ninguno
  OrdenValida = False
  
  ReDim Preserve CamposFile(100) As Campos_Tabla
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
       Line Input #NumFile, Cod_Field
       If Separador = Ninguno Then
          If InStr(Cod_Field, vbTab) > 0 Then Separador = vbTab
       End If
       Do While Len(Cod_Field) > 2
          No_Hasta = InStr(Cod_Field, Separador)
          CamposFile(CantCampos).Campo = "C" & Format$(CantCampos, "00")
          CamposFile(CantCampos).Ancho = No_Hasta
          If No_Hasta > 1 Then
             CampoTemp = TrimStrg(MidStrg(Cod_Field, 1, No_Hasta - 1))
             Select Case TextoBanco
               Case "PICHINCHA"
                  If CantCampos = 14 And TxtOrden = CampoTemp Then
                     Orden_Pago = CampoTemp  ' Orden No
                     OrdenValida = True
                  End If
                  'MsgBox CantCampos & " _ " & CampoTemp
               Case Else
                  OrdenValida = True
             End Select
             Cod_Field = TrimStrg(MidStrg(Cod_Field, No_Hasta + 1, Len(Cod_Field)))
          Else
             Cod_Field = ""
          End If
          CantCampos = CantCampos + 1
       Loop
  Close #NumFile
  
  'MsgBox "Ok"
  
  Total_Alumnos = Contador
  
  Cadena = ""
  For I = 0 To CantCampos - 1
      Cadena = Cadena & CamposFile(I).Campo & "=" & CamposFile(I).Ancho & vbCrLf
  Next I
  'MsgBox Total_Alumnos & " - " & CantCampos & " - " & OrdenValida & vbCrLf & String(100, "_") & vbCrLf & Cadena
  
 '--------------------------------------------------------
  
  Progreso_Barra.Valor_Maximo = (Total_Alumnos * 3) + 100
  
  If Not OrdenValida Then
     MsgBox "La informacion del archivo no pertenece a la Orden No. " & TxtOrden & " registrada del Banco, vuelva a seleccionar el documento correcto."
     GoTo Salida_Rutina
  End If
    
 'Procedemos a borrar los abonos recibidos
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cod_Field
      'Colocamos los datos del archivo en un array de texto
       CantCampos = 0
       Do While Len(Cod_Field) > 2
         'MsgBox UBound(CamposFile) & " - " & CantCampos & ".- " & CamposFile(35).Valor & " - " & MidStrg(Cod_Field, 1, No_Hasta - 1)
          No_Hasta = InStr(Cod_Field, Separador)
          If No_Hasta > 1 Then
             CamposFile(CantCampos).Valor = TrimStrg(MidStrg(Cod_Field, 1, No_Hasta - 1))
             Cod_Field = TrimStrg(MidStrg(Cod_Field, No_Hasta + 1, Len(Cod_Field)))
          Else
             Cod_Field = ""
          End If
          CantCampos = CantCampos + 1
       Loop
      'Procedemos a eliminar los abonos que se encuentran en el archivo, por si volvemos a subir
       Select Case TextoBanco
         Case "PICHINCHA"
              TipoDoc = CamposFile(7).Valor
              TipoProc = SinEspaciosDer(TipoDoc)
              TA.Serie = SinEspaciosDer(TrimStrg(MidStrg(TipoDoc, 1, Len(TipoDoc) - Len(TipoProc))))
              TA.Factura = Val(CamposFile(35).Valor)
              TA.Recibo_No = Format(Val(CamposFile(34).Valor), "0000000000")
              sSQL = "DELETE * " _
                   & "FROM Trans_Abonos " _
                   & "WHERE Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND TP = 'FA' " _
                   & "AND Serie = '" & TA.Serie & "' " _
                   & "AND Factura = " & TA.Factura & " " _
                   & "AND Recibo_No = '" & TA.Recibo_No & "' "
              Ejecutar_SQL_SP sSQL
       End Select
       Progreso_Barra.Mensaje_Box = "Borrando Abono No. " & TA.Recibo_No & ", documento No. " & TA.Serie & "-" & Format(TA.Factura, "000000000")
       Progreso_Esperar
    Loop
  Close #NumFile
  
  FA.Serie = Ninguno
  FA.TC = Ninguno
  FA.Factura = 0
  FA.Fecha_Desde = "01/01/2000"
  FA.Fecha_Hasta = FechaSistema
  
  Actualizar_Abonos_Facturas_SP FA
  
  AbonosAnticipados = 0
  Total_Dep_Confirmar = 0
  Trans_No = 200
  Eliminar_Asientos_SP True
  SubCtaGen = Leer_Seteos_Ctas("Cta_Anticipos_Clientes")
  Cta_Del_Banco = TrimStrg(SinEspaciosIzq(DCBanco))
  Contrato_No = Ninguno
  
  TxtFile.Text = ""
  Fecha_Tope = FechaSistema
  Total_Costo_Banco = 0
  CaptionTemp = FRecaudacionBancosCxC.Caption
  TextoImprimio = ""
  
 'Alumnos/Clientes que estan activados para Generar las Facturas
  sSQL = "SELECT Codigo, Cliente, CI_RUC, Direccion, Grupo, Email, Email2, EmailR " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY CI_RUC "
  Select_Adodc AdoClientes, sSQL
  
  FechaValida MBFechaI
  Mifecha = BuscarFecha(MBFechaI)
  FechaTexto = MBFechaI ' FechaSistema
  DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
  Label4.Caption = Dir_Dialog.Filename
  RutaGeneraFile = UCaseStrg(NombreArchivo)
  If RutaGeneraFile <> "" Then
     TotalIngreso = 0
     Contador = 0
     FileResp = 0
    'Establecemos los campos del archivo plano del Banco
     NumFile = FreeFile
     Total_Alumnos = 0
     FechaTexto = FechaSistema
     TxtFile.Text = ""
     Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          Cod_Field = Replace(Cod_Field, Chr(34), "")
          TxtFile.Text = TxtFile.Text & Cod_Field & vbCrLf
          TxtFile.SelStart = Len(TxtFile.Text)
          TxtFile.SelLength = Len(TxtFile.Text)
          TxtFile.Refresh
          
         'Comenzamos la subida de los Abonos
          CantCampos = 0
          Do While Len(Cod_Field) > 2
             No_Hasta = InStr(Cod_Field, Separador)
             CamposFile(CantCampos).Valor = TrimStrg(MidStrg(Cod_Field, 1, No_Hasta - 1))
             Cod_Field = TrimStrg(MidStrg(Cod_Field, No_Hasta + 1, Len(Cod_Field)))
             CantCampos = CantCampos + 1
          Loop
           
         'Actualizamos de que alumnos vamos a ingresar el abono
          TA.Serie = Ninguno
          TA.Factura = 0
          TA.Fecha = FechaSistema
          TA.CodigoC = Ninguno
          TA.Recibo_No = "0000000000"
          CodigoCli = Ninguno
          CodigoP = "0"
          Proceso_Ok = "PROCESO OK"
          Select Case TextoBanco
            Case "PICHINCHA"
    ''           CodigoP = CStr(Val(CamposFile(7).Valor))
    ''           FechaTexto = CamposFile(12).Valor
                 If Tipo_Carga = 1 Then
                    CodigoP = TrimStrg(CStr(Val(MidStrg(Cod_Field, 25, 19))))
                    FechaTexto = MidStrg(Cod_Field, 205, 2) & "/" & MidStrg(Cod_Field, 207, 2) & "/" & MidStrg(Cod_Field, 209, 4)
                 Else
                   'Serie de la Factura
                    TipoDoc = CamposFile(7).Valor
                    TipoProc = SinEspaciosDer(TipoDoc)
                    TipoDoc = TrimStrg(MidStrg(TipoDoc, 1, Len(TipoDoc) - Len(TipoProc)))

                    TA.Serie = SinEspaciosDer(TipoDoc)                              ' Serie
                    TA.Factura = Val(CamposFile(35).Valor)                          ' Factura
                    TA.CodigoC = CamposFile(4).Valor                                ' Codigo Cliente
                    TA.Fecha = Replace((CamposFile(25).Valor), " ", "/")            ' Fecha de Pago
                    TA.Recibo_No = Format(Val(CamposFile(34).Valor), "0000000000")  ' Recibo No
                    TA.Abono = Val(CamposFile(27).Valor)                            ' Valor del Abono

                    Proceso_Ok = CStr(TrimStrg(CamposFile(22).Valor))               ' Procesado Ok
                    If Proceso_Ok = "REVERSO OK" Then CodigoP = Format$(Val(CodigoP), "0000000000000")
                    CodigoP = TA.CodigoC
                   'Detalle del Abono
                    If TrimStrg(CamposFile(29).Valor) = "EFE" Then
                       TA.Banco = "PAGO EN EFECTIVO"
                       TA.Cheque = "VENT.: " & Replace(MidStrg(CamposFile(26).Valor, 12, 5), " ", "h") & "s"
                    Else
                       TA.Banco = "TRANS. " & CamposFile(29).Valor & "|" & CamposFile(16).Valor
                      'Tipo de Transaccion y Hora del pago
                       TA.Cheque = CamposFile(18).Valor & "-" & CamposFile(19).Valor & ": " & Replace(MidStrg(CamposFile(26).Valor, 12, 5), " ", "h") & "s"
                    End If
                 End If
            Case "BOLIVARIANO"
                  If CheqSAT.value = 1 Then
                     CodigoP = MidStrg(Cod_Field, 14, 8)
                  Else
                     CodigoP = MidStrg(Cod_Field, 1, 8)
                  End If
                  If Total_Alumnos = 0 Then
                     FechaTexto = MidStrg(Cod_Field, 12, 2) & "/" _
                                & MidStrg(Cod_Field, 10, 2) & "/" _
                                & MidStrg(Cod_Field, 6, 4)
                     CodigoP = Ninguno
                  End If
            Case "BGR_EC"
                  If Tipo_Carga = 1 Then
                     CodigoP = TrimStrg(CStr(Val(MidStrg(Cod_Field, 25, 19))))
                     FechaTexto = MidStrg(Cod_Field, 205, 2) & "/" _
                                & MidStrg(Cod_Field, 207, 2) & "/" _
                                & MidStrg(Cod_Field, 209, 4)
                  Else
                     CodigoP = CamposFile(11).Valor
                     FechaTexto = Replace(CamposFile(25).Valor, " ", "/")
                     HoraTexto = Replace(CamposFile(26).Valor, " ", ":")
                     CodigoB = CamposFile(29).Valor & ":" & CamposFile(20).Valor & "-" & Replace(CamposFile(26).Valor, " ", ":")
                  End If
            Case "INTERNACIONAL"
                  CodigoP = TrimStrg(CStr(Val(MidStrg(Cod_Field, 25, 19))))
                  FechaTexto = MidStrg(Cod_Field, 205, 2) & "/" _
                             & MidStrg(Cod_Field, 207, 2) & "/" _
                             & MidStrg(Cod_Field, 209, 4)
            Case "PACIFICO"
                  If CheqSAT.value Then
                     CodigoP = CamposFile(17).Valor
                     FechaTexto = Format$(CamposFile(11).Valor, FormatoFechas)
                  Else
                     If Total_Alumnos <> 0 Then
                        CodigoP = CamposFile(4).Valor
                        FechaTexto = MidStrg(CamposFile(6).Valor, 1, 10)
                     End If
                  End If
            Case "PRODUBANCO"
                  CodigoP = CamposFile(7).Valor
                  FechaTexto = CamposFile(12).Valor
                  CodigoB = CamposFile(14).Valor
                  NoAnio = Val(MidStrg(SinEspaciosDer(CodigoB), 1, 4))
                  If NoAnio <= "1900" And IsDate(FechaTexto) Then
                     NoMeses = Month(FechaTexto)
                     NoAnio = CStr(Year(FechaTexto))
                     Mes = MesesLetras(NoMeses)
                  End If
            Case "INTERMATICO"
                  CodigoP = CamposFile(7).Valor
                  FechaTexto = CamposFile(1).Valor
                  If Len(FechaTexto) > 10 Then
                     FechaTexto = FechaSistema
                     CodigoP = Ninguno
                  End If
                  Mifecha = FechaTexto
            Case "COOPJEP"
                  CodigoP = TrimStrg(CamposFile(16).Valor)
                  FechaTexto = CamposFile(1).Valor
            Case "CACPE"
                  CodigoP = CStr(Val(CamposFile(6).Valor))
                  FechaTexto = MidStrg(CamposFile(8).Valor, 4, 2) & "/" _
                             & MidStrg(CamposFile(8).Valor, 1, 2) & "/" _
                             & MidStrg(CamposFile(8).Valor, 7, 4)
            Case Else
                  CodigoP = Ninguno
                  TipoDoc = CamposFile(0).Valor
                  FechaTexto = CamposFile(1).Valor
                  SerieFactura = MidStrg(CamposFile(2).Valor, 1, 3) & MidStrg(CamposFile(2).Valor, 5, 3)
                  Factura_No = Val(MidStrg(CamposFile(2).Valor, 9, 10))
                  Autorizacion = CamposFile(3).Valor
                  sSQL = "SELECT F.*,C.CI_RUC " _
                       & "FROM Facturas As F, Clientes As C " _
                       & "WHERE F.Periodo = '" & Periodo_Contable & "' " _
                       & "AND F.Item = '" & NumEmpresa & "' " _
                       & "AND F.TC = '" & TipoDoc & "' " _
                       & "AND F.Serie = '" & SerieFactura & "' " _
                       & "AND F.Autorizacion = '" & Autorizacion & "' " _
                       & "AND F.Factura = " & Factura_No & " " _
                       & "AND F.CodigoC = C.Codigo "
                  Select_Adodc AdoFactura, sSQL
                  If AdoFactura.Recordset.RecordCount > 0 Then
                     CodigoP = AdoFactura.Recordset.fields("CI_RUC")
                     CodigoCli = AdoFactura.Recordset.fields("CodigoC")
                  End If
                 
          End Select
           
         'MsgBox CodigoP & vbCrLf & FechaTexto & vbCrLf & SerieFactura & vbCrLf & Factura_No
          Si_No = True
          With AdoClientes.Recordset
           If .RecordCount > 0 Then
               Do While Len(CodigoP) <= 10 And Si_No
                 .MoveFirst
                 .Find ("CI_RUC = '" & CodigoP & "' ")
                  If Not .EOF Then
                     TA.CodigoC = .fields("Codigo")
                     NombreCliente = .fields("Cliente")
                     FA.CodigoC = TA.CodigoC
                     FA.Cliente = NombreCliente
                     FA.EmailC = .fields("Email")
                     FA.EmailC2 = .fields("Email2")
                     FA.EmailR = .fields("EmailR")
                     Si_No = False
                  Else
                     CodigoP = "0" & CodigoP
                  End If
               Loop
           End If
          End With
          If Len(CodigoP) > 10 Then TA.CodigoC = Ninguno
           
          Progreso_Barra.Mensaje_Box = "Extrayendo abonos de: " & NombreCliente & ", Espere un momento"
          Progreso_Esperar
         'Procedemos a ingresar los abonos
          If TA.CodigoC <> Ninguno Then
             TotalIngreso = TotalIngreso + TA.Abono
             sSQL = "SELECT CodigoC, Cta_CxP, TC, Vencimiento, Autorizacion, Saldo_MN, Fecha " _
                  & "FROM Facturas " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND T <> 'A' " _
                  & "AND TC = 'FA' " _
                  & "AND CodigoC = '" & TA.CodigoC & "' " _
                  & "AND Serie = " & TA.Serie & " " _
                  & "AND Factura = " & TA.Factura & " " _
                  & "AND Saldo_MN > 0 "
             Select_Adodc AdoAbono, sSQL
             
             AbonosPar = NombreCliente & " (" & TA.CodigoC & "): Valor Abono: " & Format$(TA.Abono, "#,##0.00")
             
             ''If InStr(NombreCliente, "GOMEZCOELLO") > 0 Then MsgBox CodigoP & " - " & Total & ", Cliente: " & NombreCliente
             
             If AdoAbono.Recordset.RecordCount > 0 Then
                FA.Fecha = AdoAbono.Recordset.fields("Fecha")
                TA.Cta_CxP = AdoAbono.Recordset.fields("Cta_CxP")
                TA.Autorizacion = AdoAbono.Recordset.fields("Autorizacion")
                SetAdoAddNew "Trans_Abonos"
                SetAdoFields "T", Cancelado
                SetAdoFields "TP", "FA"
                SetAdoFields "CodigoC", TA.CodigoC
                SetAdoFields "Fecha", TA.Fecha
                SetAdoFields "Comprobante", "Orden No. " & Orden_Pago
                SetAdoFields "Serie", TA.Serie
                SetAdoFields "Factura", TA.Factura
                SetAdoFields "Abono", TA.Abono
                SetAdoFields "Banco", TA.Banco
                SetAdoFields "Cheque", TA.Cheque
                SetAdoFields "Cta", Cta_Del_Banco     'Cta_CajaG
                SetAdoFields "Cta_CxP", TA.Cta_CxP
                SetAdoFields "Autorizacion", TA.Autorizacion
                SetAdoFields "Recibo_No", TA.Recibo_No
                SetAdoUpdate
                
'''               'Enviar por mail el Abono receptado
'''                FA.TC = TA.TP
'''                FA.Serie = TA.Serie
'''                FA.Factura = TA.Factura
'''                FA.ClaveAcceso = FA.Autorizacion
'''                FA.Autorizacion = TA.Autorizacion
'''                FA.Fecha_C = TA.Fecha
'''                FA.Fecha_V = TA.Fecha
'''                FA.Hora_FA = TA.Cheque
'''                FA.Cliente = NombreCliente
'''                FA.Fecha_Aut = FechaSistema
'''                SRI_Autorizacion.Autorizacion = TA.Autorizacion
'''                FA.Nota = "Tipo de Abono" & vbTab & ": " & TA.Banco & vbCrLf _
'''                        & "Hora" & vbTab & vbTab & ": " & TA.Cheque & vbCrLf _
'''                        & "Documento" & vbTab & ": " & TA.Recibo_No & vbCrLf _
'''                        & "Valor Recibdo USD " & Format(TA.Abono, "#,##0.00") & vbCrLf
'''                SRI_Enviar_Mails FA, SRI_Autorizacion, "AB"

             End If
          End If
         'MsgBox NombreCliente & vbCrLf & CodigoCli & vbCrLf & CodigoP
       Loop
     Close #NumFile
     
     FA.Serie = Ninguno
     FA.TC = Ninguno
     FA.Factura = 0
     Actualizar_Abonos_Facturas_SP FA
''''Costo Bancario por deposito
''' If Costo_Banco > 0 Then
'''    Total_Costo_Banco = Total_Costo_Banco + Costo_Banco
'''    TA.T = Normal
'''    TA.TP = "CB"
'''    TA.Fecha = Mifecha
'''    TA.Cta = Cta_Del_Banco
'''    TA.Cta_CxP = Cta_Gasto_Banco
'''    TA.Banco = "COSTO BANCARIO"
'''    TA.Cheque = "DEPOSITO"
'''    TA.Factura = Factura_No
'''    TA.CodigoC = CodigoCli
'''    TA.Abono = Costo_Banco
'''    Grabar_Abonos TA
''' End If
            
 '--------------------------------------------------------------------------------------------------------
'''  Progreso_Barra.Mensaje_Box = "Procesando comprobante de abonos anticipados"
'''  Progreso_Esperar
'''
'''  Select Case TextoBanco
'''    Case "PICHINCHA"
'''    Case "BGR_EC"
'''    Case "INTERNACIONAL"
'''    Case "BOLIVARIANO": Total_Alumnos = Total_Alumnos - 1
'''    Case "PACIFICO"
'''  End Select
''' 'Realizamos un asiento de Abonos anticipado por deposito
'''  If AbonosAnticipados > 0 Then
'''     Trans_No = 200
'''     sSQL = "SELECT * " _
'''          & "FROM Asiento " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " "
'''     Select_Adodc AdoIngCaja, sSQL
'''     InsertarAsientos AdoIngCaja, Cta_Del_Banco, 0, AbonosAnticipados, 0
'''     InsertarAsientos AdoIngCaja, SubCtaGen, 0, 0, AbonosAnticipados
'''     FechaTexto = Mifecha ' FechaSistema
'''     FechaComp = Mifecha ' Este campo sirve para poder insertar donde es el comprobante en los meses
'''     RatonReloj
'''     NumComp = ReadSetDataNum("Diario", True, True)
'''     Co.TP = CompDiario
'''     Co.T = Normal
'''     Co.Fecha = FechaTexto
'''     Co.Numero = NumComp
'''     Co.Monto_Total = AbonosAnticipados
'''     Co.Concepto = "Abonos Anticipados por Depósito del día " & FechaTexto
'''     Co.CodigoB = Ninguno
'''     Co.Efectivo = AbonosAnticipados
'''     Co.Cotizacion = 0
'''     Co.Item = NumEmpresa
'''     Co.Usuario = CodigoUsuario
'''     Co.T_No = Trans_No
'''     GrabarComprobante Co
'''
'''   ' Seteamos para el siguiente comprobante
'''     RatonNormal
'''     ImprimirComprobantesDe False, Co
'''  End If
'''
''' 'MsgBox "Total_Dep_Confirmar: " & Total_Dep_Confirmar
'''  Progreso_Barra.Mensaje_Box = "Procesando comprobante de Deposito por confirmar"
'''  Progreso_Esperar
'''
'''  If Total_Dep_Confirmar > 0 Then
'''     SubCtaGen = Leer_Seteos_Ctas("Cta_Deposito_Confirmar")
'''     Trans_No = 200
'''     Eliminar_Asientos_SP True
'''     sSQL = "SELECT * " _
'''          & "FROM Asiento " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND CodigoU = '" & CodigoUsuario & "' " _
'''          & "AND T_No = " & Trans_No & " "
'''     Select_Adodc AdoIngCaja, sSQL
'''     InsertarAsientos AdoIngCaja, Cta_Del_Banco, 0, Total_Dep_Confirmar, 0
'''     InsertarAsientos AdoIngCaja, SubCtaGen, 0, 0, Total_Dep_Confirmar
'''     'MsgBox Total_Dep_Confirmar
'''     FechaTexto = Mifecha ' FechaSistema
'''     FechaComp = Mifecha ' Este campo sirve para poder insertar donde es el comprobante en los meses
'''     RatonReloj
'''     NumComp = ReadSetDataNum("Ingresos", True, True)
'''     Co.TP = CompIngreso
'''     Co.T = Normal
'''     Co.Fecha = FechaTexto
'''     Co.Numero = NumComp
'''     Co.Monto_Total = Total_Dep_Confirmar
'''     Co.Concepto = "Depósito por Confirmar del día " & FechaTexto
'''     Co.CodigoB = Ninguno
'''     Co.Efectivo = Total_Dep_Confirmar
'''     Co.Cotizacion = 0
'''     Co.Item = NumEmpresa
'''     Co.Usuario = CodigoUsuario
'''     Co.T_No = Trans_No
'''     GrabarComprobante Co
'''
'''   ' Seteamos para el siguiente comprobante
'''     RatonNormal
'''     ImprimirComprobantesDe False, Co
'''  End If
'''  Progreso_Barra.Mensaje_Box = "Actualizando Abonos de facturas"
'''  Progreso_Esperar
'''
'''  Progreso_Barra.Incremento = Contador
'''  RatonNormal
'''  FRecaudacionBancosCxC.Caption = CaptionTemp
'''  Progreso_Final
'''  Cadena = TextoImprimio
  
    MsgBox "ARCHIVO DE ABONO DEL DIA: " & FechaTexto & vbCrLf _
         & "SE ACTUALIZARON: " & Total_Alumnos & " ESTUDIANTES." & vbCrLf _
         & "EL CIERRE DIARIO DE CAJA ES POR " & Moneda & " " & Format$(TotalIngreso, "#,##0.00") & vbCrLf _
         & "EL COSTO BANCARIO ES POR " & Moneda & " " & Format$(Total_Costo_Banco, "#,##0.00") & vbCrLf _
         & "OBTENIDO DEL ARCHIVO: " & vbCrLf & RutaGeneraFile & vbCrLf
  Else
     MsgBox "No hay Archivo seleccionado"
  End If
Salida_Rutina: 'En caso de haber problemas con el archivo
End Sub

Private Sub DCTablaSRI_Change()
Dim ColorFondo As Long
  TextoBanco = Ninguno
  With AdoTablaSRI.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Descripcion = '" & DCTablaSRI.Text & "' ")
       If Not .EOF Then TextoBanco = .fields("Abreviado")
   End If
  End With
  Select Case TextoBanco
    Case "PICHINCHA":
         ColorFondo = &H80FFFF
         RutaOrigen = RutaSistema & "\LOGOS\PICHINCHA.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO DEL PICHINCHA (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case "BOLIVARIANO":
         ColorFondo = &H808000
         RutaOrigen = RutaSistema & "\LOGOS\BOLIVARIANO.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO BOLIVARIANO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = True
    Case "BGR_EC":
         ColorFondo = &H80000004
         RutaOrigen = RutaSistema & "\LOGOS\BGR_EC.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO GENERAL RUMIÑAHUI (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case "INTERNACIONAL":
         ColorFondo = &HFF8080
         RutaOrigen = RutaSistema & "\LOGOS\INTERNACIONAL.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO INTERNACIONAL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case "PACIFICO":
         ColorFondo = &HC0C000
         RutaOrigen = RutaSistema & "\LOGOS\PACIFICO.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO DEL PACIFICO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Caption = "Archivos OCP"
         CheqSAT.Visible = True
    Case "INTERMATICO"
         ColorFondo = &HC0C000
         RutaOrigen = RutaSistema & "\LOGOS\INTERMATICO.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO DEL PACIFICO: INTERMATICO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Caption = "Archivos OCP"
         CheqSAT.Visible = True
    Case "BIZBANCKPACIFICO":
         ColorFondo = &HC0C000
         RutaOrigen = RutaSistema & "\LOGOS\BIZBANCK.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO DEL PACIFICO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Caption = "Archivos OCP"
         CheqSAT.Visible = True
    Case "PRODUBANCO"
         ColorFondo = &HFFFFFF
         RutaOrigen = RutaSistema & "\LOGOS\PRODUBAN.JPG"
         FRecaudacionBancosCxC.Caption = "PRODUBANCO (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
    Case "GUAYAQUIL":
         ColorFondo = &HFF8080
         RutaOrigen = RutaSistema & "\LOGOS\GUAYAQUIL.GIF"
         FRecaudacionBancosCxC.Caption = "BANCO DE GUAYQUIL (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case "COOPJEP"
         ColorFondo = &H80FF80
         RutaOrigen = RutaSistema & "\LOGOS\COOP_JEP.GIF"
         FRecaudacionBancosCxC.Caption = "COOPERATIVA JEP (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case "CACPE"
         ColorFondo = &H80FF80
         RutaOrigen = RutaSistema & "\LOGOS\CACPE.GIF"
         FRecaudacionBancosCxC.Caption = "COOPERATIVA CACPE (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
    Case Else
         ColorFondo = &HFFFFFF
         RutaOrigen = RutaSistema & "\LOGOS\ABONOS1.GIF"
         FRecaudacionBancosCxC.Caption = "OTROS BANCOS (" & CodigoDelBanco & ")" & String(40, " ") & "EL ULTIMO CODIGO: " & Codigo
         CheqSAT.Visible = False
  End Select
  PctBanco.BackColor = ColorFondo
  Command2.BackColor = ColorFondo
  FRecaudacionBancosCxC.BackColor = ColorFondo
  PctBanco.Picture = LoadPicture(RutaOrigen)   ', 0, 0, 5000, 1100
End Sub

'''Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
'''  NombreArchivo = File1.Filename
'''  If KeyCode = vbKeyDelete Then
'''     Mensajes = "Esta seguro de Eliminar: " & File1.Filename
'''     Titulo = "Pregunta de Eliminacion"
'''     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.Filename
'''     File1.Filename = Dir1.Path & "\*.*"
'''  End If
'''
'''End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  FechaValida MBFechaF
  NuevoComp = True
  ModificarComp = False
  CopiarComp = False
  Co.CodigoB = ""
  Co.Numero = 0
  Set ftp = New cFTP
  
  Frame1.Caption = "| CUENTA A LA QUE SE VA ACREDITAR LOS ABONOS, CODIGO DEL BANCO: " & CodigoDelBanco & " |"
  Label4.Caption = Ninguno
  FRecaudacionBancosCxC.Caption = "FACTURACION DE BANCOS (" & CodigoDelBanco & ")"
  
  sSQL = "SELECT Descripcion, Abreviado, ID " _
       & "From Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "AND Abreviado <> '.' " _
       & "AND TFA <> " & Val(adFalse) & " " _
       & "ORDER BY Descripcion "
  SelectDB_Combo DCTablaSRI, AdoTablaSRI, sSQL, "Descripcion"

  sSQL = "SELECT Grupo,Count(Grupo) As Cantidad " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "AND LEN(Grupo) > 1 "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDB_Combo DCGrupoI, AdoGrupo, sSQL, "Grupo"
  SelectDB_Combo DCGrupoF, AdoGrupo, sSQL, "Grupo"
  If AdoGrupo.Recordset.RecordCount > 0 Then
     DCGrupoI.Text = AdoGrupo.Recordset.fields("Grupo")
     AdoGrupo.Recordset.MoveLast
     DCGrupoF.Text = AdoGrupo.Recordset.fields("Grupo")
  End If

  SQL2 = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo,CxC "
  Select_Adodc AdoAux, SQL2
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Cta_Cobrar = .fields("CxC")
       CxC_Clientes = .fields("Concepto")
       LogoFactura = .fields("Logo_Factura")
       AltoFactura = .fields("Largo")
       AnchoFactura = .fields("Ancho")
       EspacioFactura = .fields("Espacios")
       Pos_Factura = .fields("Pos_Factura")
       Individual = .fields("Individual")
       TipoFactura = .fields("Fact")
    End If
  End With
 'Catalogo de los Rubros ha facturar
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  Select_Adodc AdoProducto, sSQL
  
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta,Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  Cta_Del_Banco = TrimStrg(SinEspaciosIzq(DCBanco))
  Label1.BackColor = FRecaudacionBancosCxC.BackColor
  'Frame1.BackColor = FRecaudacionBancosCxC.BackColor
  'CheqBanco.BackColor = FRecaudacionBancosCxC.BackColor
  'BackColor = FRecaudacionBancosCxC.BackColor
  
  Label8.BackColor = FRecaudacionBancosCxC.BackColor
  MBFechaI.BackColor = FRecaudacionBancosCxC.BackColor
  MBFechaF.BackColor = FRecaudacionBancosCxC.BackColor
  CheqSAT.BackColor = FRecaudacionBancosCxC.BackColor
  CheqPend.BackColor = FRecaudacionBancosCxC.BackColor
  CheqMatricula.BackColor = FRecaudacionBancosCxC.BackColor
  DCBanco.BackColor = FRecaudacionBancosCxC.BackColor
  
  TxtFile.Height = MDI_Y_Max - TxtFile.Top - 50
  TxtFile.width = MDI_X_Max - 100
  Frame1.width = MDI_X_Max - Frame1.Left - 50
  DCBanco.width = Frame1.width - 200
  Label4.width = TxtFile.width
 ' ProgBarra.width = TxtFile.width
 ' ProgBarra.Top = TxtFile.Top + TxtFile.Height
'  Command2.Top = ProgBarra.Top
'  Command2.Left = ProgBarra.Left
  
  Tipo_Carga = Leer_Campo_Empresa("Tipo_Carga_Banco")
  Costo_Banco = Leer_Campo_Empresa("Costo_Bancario")
  Cta_Bancaria = Leer_Campo_Empresa("Cta_Banco")
  Cta_Gasto_Banco = Leer_Seteos_Ctas("Cta_Gasto_Bancario")
  
 'MsgBox "--->> " & Cta_Bancaria
  RatonNormal
  Cta_Banco = Ninguno
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo = '" & Cta_Del_Banco & "' ")
       If Not .EOF Then
          DCBanco = .fields("NomCuenta")
          Cta_Banco = Cta_Del_Banco
       Else
         .MoveFirst
          MsgBox "No existen cuentas asignadas o no" & vbCrLf & vbCrLf & "estan bien establecidad las cuentas contables"
          DCBanco = .fields("NomCuenta")
          Cta_Banco = SinEspaciosIzq(DCBanco)
       End If
       DCTablaSRI.SetFocus
   Else
       MsgBox "No existen cuentas asignadas o no" & vbCrLf & vbCrLf & "estan bien establecidad las cuentas contables"
       Unload FRecaudacionBancosCxC
   End If
  End With
End Sub

Private Sub Form_Deactivate()
  FRecaudacionBancosCxC.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
 'CentrarForm FRecaudacionBancosCxC
  If CodigoUsuario = "ACCESO02" Then
     Toolbar1.buttons("AlumnosContabilidad").Enabled = True
     'Command6.Visible = True
  End If
  RutaBackupXX = ""
  ConectarAdodc AdoAux
  ConectarAdodc AdoAbono
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoBanco
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoFactura
  ConectarAdodc AdoClientes
  ConectarAdodc AdoProducto
  ConectarAdodc AdoTablaSRI
  ConectarAdodc AdoPendiente
  ConectarAdodc AdoIngCaja
  RutaBackup = RutaSysBases & "\BANCO"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Detalle = 'Deuda Pendiente' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     AdoAux.Recordset.fields("Fecha_Inicial") = MBFechaF
     AdoAux.Recordset.fields("Fecha_Final") = MBFechaF
     AdoAux.Recordset.Update
  End If
End Sub

Public Sub Generar_Bolivariano()
'022750270
'MsgBox RutaSysBases
RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\ALUMNOS" & CodigoDelBanco & ".TXT")
NumFileAbonos = FreeFile
NumFileDetalle = FreeFile
NumFileFacturas = FreeFile
NumFileAlumnos = FreeFile
TipoDoc = "0"
If CheqMatricula.value = 1 Then TipoDoc = "1"

Contador = 0
 
'''ProgBarra.value = 0
''ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaF.Text)
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoAux.Recordset
 If .RecordCount > 0 Then
    'Cabecera
     Print #NumFileFacturas, "999";
     Print #NumFileFacturas, CodigoDelBanco;
     Print #NumFileFacturas, TipoDoc;
     Print #NumFileFacturas, Space(11);
     Print #NumFileFacturas, Mifecha
    .MoveFirst
     ''ProgBarra.Max = .RecordCount
     'Trama / Detalle
     Do While Not .EOF
        FRecaudacionBancosCxC.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        SaldoPendiente = 0
        Total_Factura = 0
        Monto_Total = 0
        Total = 0
        ''ProgBarra.value = Contador
        CodigoCli = .fields("CI_RUC")
        Codigo = "0"
        For I = 1 To Len(.fields("CI_RUC"))
            If IsNumeric(MidStrg(.fields("CI_RUC"), I, 1)) Then Codigo = Codigo & MidStrg(.fields("CI_RUC"), I, 1)
        Next I
        Codigo = TrimStrg(Str(Val(Codigo)))
        Codigo = Codigo & String(15 - Len(Codigo), " ")
      ' MsgBox "|" & Codigo & "|"
        NombreCliente = SetearBlancos(MidStrg(.fields("Cliente"), 1, 30), 30, 0, False)
        Codigo1 = TrimStrg(MidStrg(SinEspaciosIzq(.fields("Direccion")), 1, 15))
        Codigo3 = TrimStrg(MidStrg(SinEspaciosDer(.fields("Direccion")), 1, 3))
        Codigo2 = TrimStrg(MidStrg(.fields("Direccion"), Len(Codigo1) + 1, Len(.fields("Direccion"))))
        Codigo4 = MidStrg(.fields("Casilla"), 1, 10)
        Saldo_ME = 0: Total_Desc = 0: SaldoPendiente = 0
        If AdoPendiente.Recordset.RecordCount > 0 Then
           AdoPendiente.Recordset.MoveFirst
           AdoPendiente.Recordset.Find ("CI_RUC Like '" & CodigoCli & "' ")
           If Not AdoPendiente.Recordset.EOF Then SaldoPendiente = AdoPendiente.Recordset.fields("Saldo_Pend")
        End If
        If CheqPend.value = 1 Then SaldoPendiente = .fields("Total_MN")
        Total_Factura = .fields("Total_MN")
        Monto_Total = Total_Factura
        Total = SaldoPendiente
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
        Codigo2 = TrimStrg(MidStrg(Codigo2, 1, Len(Codigo2) - Len(SinEspaciosDer(Codigo2))))
        Codigo1 = SetearBlancos(Codigo1, 15, 0, False)
        Codigo2 = SetearBlancos(Codigo2, 15, 0, False)
        Codigo3 = SetearBlancos(Codigo3, 3, 0, False)
        Codigo4 = SetearBlancos(Codigo4, 10, 0, False)
        If TrimStrg(Codigo4) = Ninguno Then Codigo4 = String(10, " ")
      ' Total = Total - Monto_Total
        If Total < 0 Then Total = 0
      ' Empieza la trama por Alumno
        'MsgBox NombreCliente & vbCrLf & Total
        Print #NumFileFacturas, CodigoDelBanco;                        ' Colegio/Institucion
        Print #NumFileFacturas, Codigo;                                ' Codigo Alumno
        Print #NumFileFacturas, Mifecha;                               ' Fecha Pen: FechaTexto = FechaTexto1
        Print #NumFileFacturas, TipoDoc & "  ";                        ' Proceso
        Print #NumFileFacturas, Format$(Total, "00000000.00");          ' Valor
        Print #NumFileFacturas, FechaTexto;                            ' Fecha Cobis
        Print #NumFileFacturas, "01/01/1900";                          ' Fecha Pago "01/01/1900";
        Print #NumFileFacturas, "N";                                   ' Estado = N
        Print #NumFileFacturas, Sin_Signos_Especiales(NombreCliente); ' Nombre Alumno
        Print #NumFileFacturas, Codigo2;                               ' Nombre del Curso
        Print #NumFileFacturas, Codigo3;                               ' Nombre del Paralelo
        Print #NumFileFacturas, Codigo1;                               ' Nombre de la Seccion
        Print #NumFileFacturas, Format$(Monto_Total, "00000000.00");    ' Valor Mes
        Print #NumFileFacturas, Codigo4;                               ' Pago por Deposito de Cuenta
        Print #NumFileFacturas, "1";                                   ' Moneda = 1
        Print #NumFileFacturas, Format$(Total, "00000000.00");          ' Valor 2
        Print #NumFileFacturas, Format$(Total, "00000000.00")           ' Valor 1
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
'''ProgBarra.value = 'ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Public Sub Generar_Coop_Jep()
  RatonReloj
 'Facturas Emitidas del mes
  sSQL = "SELECT CI_RUC As 'CODIGO ALUMNO', C.Cliente As 'NOMBRE ALUMNO', C.Direccion As 'CURSO', "
  If CheqMatricula.value = 1 Then
     sSQL = sSQL & "SUM(Saldo_MN) As 'MATRICULA', '0' As 'PENSION', "
  Else
     sSQL = sSQL & "'0' As 'MATRICULA',SUM(Saldo_MN) As 'PENSION', "
  End If
  sSQL = sSQL & "'' As 'TRANSPORTE','' As 'REFRIGERIO','' As 'DERECHOS DE EXAMEN', " _
       & "'' As 'DEUDA PENDIENTE','' As 'AGENDA','' As 'RECARGOS','' As 'TALLERES SEMINARIOS','' As 'OTROS', " _
       & "SUM(Saldo_MN) As 'VALOR TOTAL',SUM(Con_IVA) As 'BICONIVA','' As 'ICE',SUM(IVA) As 'IVA'," _
       & "SUM(Sin_IVA) As 'BISINIVA','' As 'BI NO OBJETO IVA',C.Email As 'MAIL' " _
       & "FROM Facturas As F,Clientes As C, Clientes_Matriculas AS CM " _
       & "WHERE F.Fecha BETWEEN #" & BuscarFecha(MBFechaI) & "# and #" & BuscarFecha(MBFechaF) & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T = 'P' " _
       & "AND NOT F.TC IN ('C','P') " _
       & "AND F.CodigoC = C.Codigo " _
       & "AND F.CodigoC = CM.Codigo " _
       & "AND F.Periodo = CM.Periodo " _
       & "AND F.Item = CM.Item " _
       & "GROUP BY C.CI_RUC,C.Cliente,C.Direccion,CM.TD,CM.Cedula_R,CM.Representante,C.Celular,C.Email " _
       & "HAVING SUM(Saldo_MN) > 0 " _
       & "ORDER BY C.Direccion,C.Cliente,C.CI_RUC "
  Select_Adodc AdoFactura, sSQL
RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\CARGA_JEP_" & Replace(FechaSistema, "/", "-") & ".TXT")
NumFileFacturas = FreeFile
TipoDoc = "0"
Contador = 0
''ProgBarra.value = 0
'ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaF.Text)
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
     'ProgBarra.Max = .RecordCount
     'Trama / Detalle
     Print #NumFileFacturas, "CODIGO ALUMNO" & vbTab;                        ' Colegio/Institucion
     Print #NumFileFacturas, "NOMBRE ALUMNO" & vbTab;                                ' Codigo Alumno
     Print #NumFileFacturas, "CURSO" & vbTab;                           ' Fecha Pen: FechaTexto = FechaTexto1
     Print #NumFileFacturas, "MATRICULA" & vbTab;
     Print #NumFileFacturas, "PENSION" & vbTab;
     Print #NumFileFacturas, "TRANSPORTE" & vbTab;
     Print #NumFileFacturas, "REFRIGERIO" & vbTab;
     Print #NumFileFacturas, "DERECHOS DE EXAMEN" & vbTab;
     Print #NumFileFacturas, "DEUDA PENDIENTE" & vbTab;
     Print #NumFileFacturas, "AGENDA" & vbTab;
     Print #NumFileFacturas, "RECARGOS" & vbTab;
     Print #NumFileFacturas, "TALLERES SEMINARIOS" & vbTab;
     Print #NumFileFacturas, "OTROS" & vbTab;
     Print #NumFileFacturas, "VALOR TOTAL" & vbTab;                         ' Fecha Pago "01/01/1900";
     Print #NumFileFacturas, "BICONIVA" & vbTab;
     Print #NumFileFacturas, "ICE" & vbTab;
     Print #NumFileFacturas, "IVA" & vbTab;
     Print #NumFileFacturas, "BISINIVA" & vbTab;
     Print #NumFileFacturas, "BI NO OBJETO IVA" & vbTab;
     Print #NumFileFacturas, "MAIL"
     Do While Not .EOF
        Codigo = .fields("MAIL")
        If Len(Codigo) = 1 Then Codigo = ""
        Print #NumFileFacturas, .fields("CODIGO ALUMNO") & vbTab;                        ' Colegio/Institucion
        Print #NumFileFacturas, .fields("NOMBRE ALUMNO") & vbTab;                                ' Codigo Alumno
        Print #NumFileFacturas, .fields("CURSO") & vbTab;                           ' Fecha Pen: FechaTexto = FechaTexto1
        Print #NumFileFacturas, Campo_Blanco(.fields("MATRICULA")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("PENSION")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("TRANSPORTE")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("REFRIGERIO")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("DERECHOS DE EXAMEN")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("DEUDA PENDIENTE")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("AGENDA")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("RECARGOS")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("TALLERES SEMINARIOS")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("OTROS")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("VALOR TOTAL")) & vbTab;                         ' Fecha Pago "01/01/1900";
        Print #NumFileFacturas, Campo_Blanco(.fields("BICONIVA")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("ICE")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("IVA")) & vbTab;
        Print #NumFileFacturas, Campo_Blanco(.fields("BISINIVA")) & vbTab;
        If Val(.fields("BISINIVA")) > 0 Then
           Print #NumFileFacturas, Campo_Blanco(.fields("BISINIVA")) & vbTab;
        Else
           Print #NumFileFacturas, Campo_Blanco(.fields("BICONIVA")) & vbTab;
        End If
        Print #NumFileFacturas, Codigo
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
''ProgBarra.value = 'ProgBarra.Max
RatonNormal
GenerarDataTexto FRecaudacionBancosCxC, AdoFactura
MsgBox "Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Public Sub Actualizar_Bolivariano()
'MsgBox RutaSysBases
RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\ACTALUMNOS" & CodigoDelBanco & ".TXT")
NumFileFacturas = FreeFile
TipoDoc = "0"
If CheqMatricula.value = 1 Then TipoDoc = "1"
Contador = 0
''ProgBarra.value = 0
'ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaF.Text)
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoFactura.Recordset
 If .RecordCount > 0 Then
    'Cabecera
     Print #NumFileFacturas, "999";
     Print #NumFileFacturas, CodigoDelBanco;
     Print #NumFileFacturas, TipoDoc;
     Print #NumFileFacturas, Space(11);
     Print #NumFileFacturas, FechaTexto1
    .MoveFirst
     'ProgBarra.Max = .RecordCount
     'Trama / Detalle
     Do While Not .EOF
        'FRecaudacionBancosCxC.Caption = .Fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        SaldoPendiente = 0
        Total_Factura = 0
        Monto_Total = 0
        Total = 0
        'ProgBarra.value = Contador
        CodigoCli = .fields("CI_RUC")
        Codigo = "0"
        For I = 1 To Len(.fields("CI_RUC"))
            If IsNumeric(MidStrg(.fields("CI_RUC"), I, 1)) Then Codigo = Codigo & MidStrg(.fields("CI_RUC"), I, 1)
        Next I
        Codigo = TrimStrg(Str(Val(Codigo)))
        Codigo = Codigo & String(15 - Len(Codigo), " ")
      ' MsgBox "|" & Codigo & "|"
        NombreCliente = SetearBlancos(MidStrg(.fields("Cliente"), 1, 30), 30, 0, False)
        Codigo1 = TrimStrg(MidStrg(SinEspaciosIzq(.fields("Direccion")), 1, 15))
        Codigo3 = TrimStrg(MidStrg(SinEspaciosDer(.fields("Direccion")), 1, 3))
        Codigo2 = TrimStrg(MidStrg(.fields("Direccion"), Len(Codigo1) + 1, Len(.fields("Direccion"))))
        Codigo4 = MidStrg(.fields("Casilla"), 1, 10)
        Total_Factura = .fields("Abonos")
        Monto_Total = Total_Factura
        If Codigo1 = "" Then Codigo1 = Ninguno
        If Codigo2 = "" Then Codigo2 = Ninguno
        If Codigo3 = "" Then Codigo3 = Ninguno
        Codigo2 = TrimStrg(MidStrg(Codigo2, 1, Len(Codigo2) - Len(SinEspaciosDer(Codigo2))))
        Codigo1 = SetearBlancos(Codigo1, 15, 0, False)
        Codigo2 = SetearBlancos(Codigo2, 15, 0, False)
        Codigo3 = SetearBlancos(Codigo3, 3, 0, False)
        Codigo4 = SetearBlancos(Codigo4, 10, 0, False)
        If TrimStrg(Codigo4) = Ninguno Then Codigo4 = String(10, " ")
      ' Total = Total - Monto_Total
        If Total < 0 Then Total = 0
      ' Empieza la trama por Alumno
        'MsgBox NombreCliente & vbCrLf & Total
        Print #NumFileFacturas, CodigoDelBanco;                        ' Colegio/Institucion
        Print #NumFileFacturas, Codigo;                                ' Codigo Alumno
        Print #NumFileFacturas, FechaTexto1;                           ' Fecha Pen: FechaTexto = FechaTexto1
        Print #NumFileFacturas, TipoDoc & "  ";                        ' Proceso
        Print #NumFileFacturas, Format$(Total_Factura, "00000000.00");  ' Valor
        Print #NumFileFacturas, FechaTexto;                            ' Fecha Cobis
        Print #NumFileFacturas, "01/01/1900";                          ' Fecha Pago "01/01/1900";
        Print #NumFileFacturas, "N";                                   ' Estado = N
        Print #NumFileFacturas, Sin_Signos_Especiales(NombreCliente); ' Nombre Alumno
        Print #NumFileFacturas, Codigo2;                               ' Nombre del Curso
        Print #NumFileFacturas, Codigo3;                               ' Nombre del Paralelo
        Print #NumFileFacturas, Codigo1;                               ' Nombre de la Seccion
        Print #NumFileFacturas, Format$(Total_Factura, "00000000.00");    ' Valor Mes
        Print #NumFileFacturas, Codigo4;                               ' Pago por Deposito de Cuenta
        Print #NumFileFacturas, "1";                                   ' Moneda = 1
        Print #NumFileFacturas, Format$(Total_Factura, "00000000.00");  ' Valor 2
        Print #NumFileFacturas, Format$(Total_Factura, "00000000.00");  ' Valor 1
        Print #NumFileFacturas, Space(97) & .fields("Comprobante")
        Contador = Contador + 1
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
''ProgBarra.value = 'ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile
End Sub

Public Sub Generar_Produbanco()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Fecha_Meses As String
Dim ValorStr As String
Dim CamposFile() As Campos_Tabla
Dim Total_Banco As Currency
    
  Total_Banco = 0
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  NumFileFacturas = FreeFile
  Fecha_Meses = MBFechaI & " al " & MBFechaF
  Fecha_Meses = Replace(Fecha_Meses, "/", "-")
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\PRODUBANCO RECAUDACION " & Fecha_Meses & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoDetalle.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("CodigoC")
          NombreCliente = TrimStrg(MidStrg(Sin_Signos_Especiales(.fields("Cliente")), 1, 40))
          CodigoP = .fields("CI_RUC")
          CodigoC = .fields("CI_RUC")
          Saldo = .fields("Total") - .fields("Total_Desc")
          Total = Saldo
          DireccionCli = .fields("Direccion")
          GrupoNo = .fields("Grupo")
          Detalle = GrupoNo & "-" & .fields("Producto") & "-" & .fields("Mes")
         'MsgBox NombreCliente
          Contador = Contador + 1
          ValorStr = CStr(Saldo * 100)
          ValorStr = String(13 - Len(ValorStr), "0") & ValorStr
         'MsgBox ValorStr
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          Codigo3 = SinEspaciosDer(DireccionCli)
          'DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          If Len(Cta_Bancaria) < 10 Then Cta_Bancaria = String(10 - Len(Cta_Bancaria), "0") & Cta_Bancaria
          Codigo4 = .fields("Grupo") & " De: " & .fields("Ticket") & "-" & MesesLetras(.fields("Mes_No"))
          Print #NumFileFacturas, "CO" & vbTab;            '01
          Print #NumFileFacturas, Cta_Bancaria & vbTab;    '02
          Print #NumFileFacturas, Contador & vbTab;        '03
          Print #NumFileFacturas, vbTab;                   '04
          Print #NumFileFacturas, CodigoP & vbTab;         '05
          Print #NumFileFacturas, "USD" & vbTab;           '06
          Print #NumFileFacturas, ValorStr & vbTab;        '07
          Print #NumFileFacturas, "REC" & vbTab;           '08
          Print #NumFileFacturas, CodigoDelBanco & vbTab;  '09
          Print #NumFileFacturas, vbTab;                   '10
          Print #NumFileFacturas, vbTab;                   '11
          Print #NumFileFacturas, "R" & vbTab;             '12 - C
          Print #NumFileFacturas, RUC & vbTab;             '13 - CodigoP
          Print #NumFileFacturas, NombreCliente & vbTab;   '14
          Print #NumFileFacturas, vbTab;                   '15
          Print #NumFileFacturas, vbTab;                   '16
          Print #NumFileFacturas, vbTab;                   '17
          Print #NumFileFacturas, vbTab;                   '18
          Print #NumFileFacturas, Codigo4 & vbTab;         '19
          Print #NumFileFacturas, vbTab;                   '20
          Print #NumFileFacturas, vbTab                    '21
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
 ' ProgBarra.value = ProgBarra.Max
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
'  LabelAbonos.Caption = Format$(Total_Banco, "#,##0.00")
  MsgBox "SE HA GENERADO EL SIGUIENTE ARCHIVO:" & vbCrLf & vbCrLf _
       & RutaGeneraFile & vbCrLf
End Sub

Public Sub Generar_Internacional()
Dim AuxNumEmp As String
Dim Traza As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
Dim Separador As String
Dim TRep As Tipo_Beneficiarios
  RatonReloj
  Separador = vbTab
  If MBFechaI <> MBFechaF Then
     Traza = Replace(MBFechaI, "/", "-") & "_al_" & Replace(MBFechaF, "/", "-")
  Else
     Traza = Replace(MBFechaI, "/", "-")
  End If
  
  RutaGeneraFileFacturas = RutaSysBases & "\BANCO\FACTURAS\CxC_MES_" & Traza & ".TXT"
  Traza = ""
 'Abrimo los archivo que vamos ha necesitar
  NumFileFacturas = FreeFile
  Open RutaGeneraFileFacturas For Output As #NumFileFacturas
  FechaTexto = FechaSistema
  Mifecha = BuscarFecha(MBFechaI)
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 1
  Total = 0
' Comenzamos a generar el archivo: COALU.TXT
  With AdoDetalle.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("CodigoC")
          TRep = Leer_Datos_Clientes(CodigoCli)
          NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 40))
          CodigoB = Format$(.fields("CI_RUC"), "00000000")
          Abono = .fields("Total") - .fields("Total_Desc") - .fields("Total_Desc2")
          If Abono <= 0 Then Abono = 0
          DireccionCli = .fields("Direccion")
          GrupoNo = .fields("Grupo")
          Detalle = GrupoNo & "-" & .fields("Producto") & "-" & .fields("Mes")
         'Empieza la trama
          Traza = "CO" & Separador _
                & TrimStrg(MidStrg(Cta_Bancaria, 1, 20)) & Separador _
                & Contador & Separador _
                & Contador & Separador _
                & CodigoB & Separador _
                & "USD" & Separador _
                & CStr(Abono * 100) & Separador _
                & "REC" & Separador _
                & "32" & Separador _
                & Separador _
                & Separador _
                & TRep.TD_Rep & Separador _
                & TRep.RUC_CI_Rep & Separador _
                & NombreCliente & Separador _
                & Separador _
                & Separador _
                & Separador _
                & Separador _
                & Detalle & String(13, vbTab)    ' 19
          If Abono > 0 Then
             Print #NumFileFacturas, Traza
             Contador = Contador + 1
          End If
          FRecaudacionBancosCxC.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
         .MoveNext
       Loop
   End If
  End With
 'Finalizamos los Archivos
  Close #NumFileFacturas
  
  Separador = vbTab
  If MBFechaI <> MBFechaF Then
     Traza = Replace(MBFechaI, "/", "-") & "_al_" & Replace(MBFechaF, "/", "-")
  Else
     Traza = Replace(MBFechaI, "/", "-")
  End If
  RutaGeneraFileFacturas = RutaSysBases & "\BANCO\FACTURAS\CxC_Pendiente_" & Traza & ".TXT"
  Traza = ""
 'Abrimo los archivo que vamos ha necesitar
  NumFileFacturas = FreeFile
  Open RutaGeneraFileFacturas For Output As #NumFileFacturas
  FechaTexto = FechaSistema
  Mifecha = BuscarFecha(MBFechaI)
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  Contador = 1
  Total = 0
' Comenzamos a generar el archivo: COALU.TXT
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .fields("CodigoC")
          TRep = Leer_Datos_Clientes(CodigoCli)
          NombreCliente = TrimStrg(MidStrg(.fields("Cliente"), 1, 40))
          CodigoB = Format$(.fields("CI_RUC"), "00000000")
          Abono = .fields("Saldo_Pend")
          If Abono <= 0 Then Abono = 0
          DireccionCli = .fields("Direccion")
          GrupoNo = .fields("Grupo")
          Detalle = GrupoNo & "-Saldo Pendiente-00"
         'Empieza la trama
          Traza = "CO" & Separador _
                & TrimStrg(MidStrg(Cta_Bancaria, 1, 20)) & Separador _
                & Contador & Separador _
                & Contador & Separador _
                & CodigoB & Separador _
                & "USD" & Separador _
                & CStr(Abono * 100) & Separador _
                & "REC" & Separador _
                & "32" & Separador _
                & Separador _
                & Separador _
                & TRep.TD_Rep & Separador _
                & TRep.RUC_CI_Rep & Separador _
                & NombreCliente & Separador _
                & Separador _
                & Separador _
                & Separador _
                & Separador _
                & Detalle & String(13, vbTab)    ' 19
          If Abono > 0 Then
             Print #NumFileFacturas, Traza
             Contador = Contador + 1
          End If
          FRecaudacionBancosCxC.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
         .MoveNext
       Loop
   End If
  End With
 'Finalizamos los Archivos
  Close #NumFileFacturas
 '''ProgBarra.value = 'ProgBarra.Max
  RatonNormal
  If MBFechaI <> MBFechaF Then
     Traza = Replace(MBFechaI, "/", "-") & "_al_" & Replace(MBFechaF, "/", "-")
  Else
     Traza = Replace(MBFechaI, "/", "-")
  End If
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS: " & vbCrLf & vbCrLf _
       & RutaSysBases & "\BANCO\FACTURAS\CxC_MES_" & Traza & ".TXT" _
       & vbCrLf & vbCrLf _
       & RutaSysBases & "\BANCO\FACTURAS\CxC_Pendiente_" & Traza & ".TXT"
End Sub

Public Sub Generar_Pichincha()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
  
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  SumaBancos = 0
  NumFileFacturas = FreeFile
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCREC" & Month(MBFechaI) & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoPendiente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("CodigoC")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          'MsgBox NombreCliente
          'Factura_No = Factura_No + 1
          Factura_No = .fields("Factura")
          SerieFactura = .fields("Serie")
          Total = .fields("Saldo_MN")
          Saldo = .fields("Saldo_MN") * 100
          CodigoP = .fields("CI_RUC")
          CodigoC = CStr(Val(.fields("CI_RUC")))
          CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = .fields("Anio_Mes")
          If Len(.fields("Actividad")) < 3 Then
             If Tipo_Carga = 1 Then
               'Tipo Gualaceo
                Print #NumFileFacturas, "CO" & vbTab;
                Print #NumFileFacturas, CodigoC & vbTab;
                Print #NumFileFacturas, "USD" & vbTab;
                Print #NumFileFacturas, Saldo & vbTab;
                Print #NumFileFacturas, "REC" & vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                If CheqMatricula.value = 1 Then
                   Codigo4 = "MATRICULAS DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
                Else
                   Codigo4 = "PENSION ACUMULADA DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
                End If
                Print #NumFileFacturas, UCaseStrg(Codigo4) & vbTab;
                Print #NumFileFacturas, "N" & vbTab;
                Print #NumFileFacturas, Format$(Val(CodigoC), "0000000000") & vbTab;
                Print #NumFileFacturas, MidStrg(NombreCliente, 1, 40) & vbTab
             Else
               'Tipo General
                Print #NumFileFacturas, "CO" & vbTab;
                Print #NumFileFacturas, Cta_Bancaria & vbTab;
                Print #NumFileFacturas, Contador & vbTab;
                Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;
                Print #NumFileFacturas, CodigoP & vbTab;
                Print #NumFileFacturas, "USD" & vbTab;
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;
                Print #NumFileFacturas, "REC" & vbTab;
                Print #NumFileFacturas, "10" & vbTab;
                Print #NumFileFacturas, vbTab;           'CTE/AHO
                Print #NumFileFacturas, "0" & vbTab;     'No. Cta Cte/Aho
                Print #NumFileFacturas, "R" & vbTab;
                Print #NumFileFacturas, RUC & vbTab;
                Print #NumFileFacturas, MidStrg(Codigo1, 6, 2) & " " & MidStrg(NombreCliente, 1, 37) & vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, vbTab;
                Print #NumFileFacturas, MidStrg(Codigo1, 6, 2) & vbTab;
                'Print #NumFileFacturas, "Pensión Acumulada" & vbTab;
                If CheqMatricula.value = 1 Then
                   Codigo4 = .fields("Grupo") & " Matricula "
                Else
                   Codigo4 = .fields("Grupo") & " Pension "
                End If
                Codigo4 = Codigo4 & String(26 - Len(Codigo4), " ") & SerieFactura & " " & Codigo1
                Print #NumFileFacturas, Codigo4 & vbTab;
                Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab
                SumaBancos = SumaBancos + .fields("Saldo_MN")
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
' Comenzamos a generar el archivo: SCCOB.TXT
  Mes = Month(MBFechaI)
  Anio = Val(MidStrg(Format$(Year(MBFechaI), "0000"), 2, 3))
  Dia = "15"
  NumFileFacturas = FreeFile
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCCOB" & Month(MBFechaI) & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoPendiente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("CodigoC")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          Factura_No = Factura_No + 1
          Total = .fields("Saldo_MN")
          Saldo = .fields("Saldo_MN") * 100
          CodigoP = .fields("CI_RUC")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          If Len(.fields("Actividad")) >= 3 Then
             Print #NumFileFacturas, "CO" & vbTab;
             Print #NumFileFacturas, Cta_Bancaria & vbTab;
             Print #NumFileFacturas, Contador & vbTab;
             Print #NumFileFacturas, Format$(Factura_No, "0000000000") & vbTab;
             Print #NumFileFacturas, CodigoP & vbTab;
             Print #NumFileFacturas, "USD" & vbTab;
             Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab;
             Print #NumFileFacturas, "CTA" & vbTab;
             Print #NumFileFacturas, "10" & vbTab;
             NumStrg = SinEspaciosIzq(.fields("Actividad"))
             If Len(NumStrg) = 3 Then
                Print #NumFileFacturas, SinEspaciosIzq(.fields("Actividad")) & vbTab;      'CTE/AHO
                Print #NumFileFacturas, SinEspaciosDer(.fields("Actividad")) & vbTab;     'No. Cta Cte/Aho
             Else
                Print #NumFileFacturas, vbTab;      'CTE/AHO
                Print #NumFileFacturas, vbTab;      'No. Cta Cte/Aho
             End If
             Print #NumFileFacturas, "R" & vbTab;
             Print #NumFileFacturas, RUC & vbTab;
             Print #NumFileFacturas, MidStrg(Month(MBFechaI) & " " & NombreCliente, 1, 40) & vbTab;
             Print #NumFileFacturas, vbTab;
             Print #NumFileFacturas, vbTab;
             Print #NumFileFacturas, vbTab;
             Print #NumFileFacturas, vbTab;
             Print #NumFileFacturas, Month(MBFechaI) & vbTab;
             Print #NumFileFacturas, "Pensión Acumulada" & vbTab;
             Print #NumFileFacturas, Format$(Saldo, "0000000000000") & vbTab
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  ''ProgBarra.value = 'ProgBarra.Max
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS:" & vbCrLf & vbCrLf _
       & UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCREC" & Month(MBFechaI) & ".TXT") & vbCrLf & vbCrLf _
       & UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\SCCOB" & Month(MBFechaI) & ".TXT") & vbCrLf & vbCrLf _
       & "Valor Total a Recaudar USD " & Format(SumaBancos, "#,##0.00")
End Sub

Public Sub Generar_BGR_EC()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Tipo_Carga As Byte
Dim CamposFile() As Campos_Tabla
  Mifecha = BuscarFecha(MBFechaI)
  MiMes = Format$(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))

  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
' Comenzamos a generar el archivo: BGR_MES_XX
  EsComa = True
  Estab = False
  Contador = 0
  Factura_No = 0
  NumFileFacturas = FreeFile
  RutaGeneraFile = UCaseStrg(RutaSysBases & "\BANCO\FACTURAS\BGR_MES_" & Format$(Month(MBFechaI), "00") & ".TXT")
  Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
  With AdoAux.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .fields("CodigoC")
          NombreCliente = Sin_Signos_Especiales(.fields("Cliente"))
          Factura_No = Factura_No + 1
          SaldoPendiente = 0
          'MsgBox AdoPendiente.Recordset.RecordCount
          If AdoPendiente.Recordset.RecordCount > 0 Then
             AdoPendiente.Recordset.MoveFirst
             AdoPendiente.Recordset.Find ("CodigoC Like '" & CodigoCli & "' ")
             If Not AdoPendiente.Recordset.EOF Then SaldoPendiente = AdoPendiente.Recordset.fields("Saldo_Pend")
          End If
          If CheqPend.value = 1 Then
             Total = .fields("Saldo_MN")
          Else
             Total = SaldoPendiente
          End If
          Saldo = Total * 100
          'If Saldo > 0 Then MsgBox Saldo
          CodigoP = .fields("CI_RUC")
          CodigoC = .fields("CI_RUC")
          ' CodigoC = CodigoC & String(Abs(4 - Len(CodigoC)), " ")
          DireccionCli = Sin_Signos_Especiales(.fields("Direccion"))
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = TrimStrg(MidStrg(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = TrimStrg(SinEspaciosDer(DireccionCli))
          Codigo1 = Format$(MidStrg(GrupoNo, 1, 1), "00")
          If Saldo > 0 Then
            'Tipo trama
             Print #NumFileFacturas, "CO" & vbTab;
             Print #NumFileFacturas, CodigoC & vbTab;
             Print #NumFileFacturas, "USD" & vbTab;
             Print #NumFileFacturas, Saldo & vbTab;
             Print #NumFileFacturas, "REC" & vbTab;
             Print #NumFileFacturas, vbTab;
             Print #NumFileFacturas, vbTab;
             If CheqMatricula.value = 1 Then
                Codigo4 = "MATRICULAS DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
             Else
                Codigo4 = "PENSION DE " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & "-" & Year(MBFechaI)
             End If
             Print #NumFileFacturas, UCaseStrg(TrimStrg(MidStrg(Codigo4, 1, 40))) & vbTab;
             Print #NumFileFacturas, "N" & vbTab;
             Print #NumFileFacturas, Format$(Val(CodigoC), "0000000000") & vbTab;
             Print #NumFileFacturas, MidStrg(NombreCliente, 1, 40) & vbTab
          End If
         .MoveNext
       Loop
   End If
  End With
  Close #NumFileFacturas
  RatonNormal
  ''ProgBarra.value = 'ProgBarra.Max
 'FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  MsgBox "SE GENERARON LOS SIGUIENTES ARCHIVOS:" & vbCrLf & vbCrLf _
       & RutaGeneraFile & vbCrLf & vbCrLf
End Sub

Public Sub Generar_Pacifico()
Dim RutaGeneraFile1 As String
Dim RutaGeneraFile2 As String
Dim FormCaption As String
'022750270
'MsgBox RutaSysBases
FormCaption = FRecaudacionBancosCxC.Caption
RutaGeneraFile1 = RutaSysBases & "\BANCO\FACTURAS\BIZBANK CODIGO DEL " _
                & Replace(MBFechaI, "/", "-") & " AL " & Replace(MBFechaF, "/", "-") & ".TXT"
RutaGeneraFile1 = UCaseStrg(RutaGeneraFile1)
NumFileFacturas = FreeFile
TipoDoc = "0"
Contador = 0
''ProgBarra.value = 0
'ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
TxtFile = ""
Open RutaGeneraFile1 For Output As #NumFileFacturas ' Abre el archivo.
With AdoPendiente.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TxtFile = "TOTAL NOMINA DE RECAUDACION:" & vbCrLf
     'ProgBarra.Max = (.RecordCount * 2) + 10
     Codigo3 = TrimStrg(MidStrg(NombreEmpresa, 1, 30))
     Total = 0
     TotalIngreso = 0
     Total_Factura = 0
     IE = 1
     JE = 1
     KE = 1
     Grupo_No = .fields("Grupo")
     Codigo = .fields("CodigoC")
     Do While Not .EOF
        FRecaudacionBancosCxC.Caption = "Por Código: " & .fields("Grupo") & " - " & Format$(Contador / (.RecordCount * 2), "00%")
        Contador = Contador + 1
        'ProgBarra.value = Contador
        CodigoCli = .fields("CI_RUC")
        NombreCliente = Sin_Signos_Especiales(TrimStrg(MidStrg(.fields("Cliente"), 1, 30)))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & " " & Year(MBFechaI)
        Codigo2 = UCaseStrg(.fields("Grupo") & " " & Replace(MBFechaF, "/", " "))
        Total_Factura = .fields("Saldo_Pend")
        If Costo_Banco > 0 Then Total_Factura = Total_Factura + Costo_Banco
        Total = Total + Total_Factura
        TotalIngreso = TotalIngreso + Total_Factura
        I = Int(Total_Factura)
        J = (Total_Factura - Int(Total_Factura)) * 100
        'MsgBox Grupo_No & "(" & JE & ")" & vbCrLf & Total_Factura & vbCrLf & Total & vbCrLf & TotalIngreso
        If Len(.fields("Actividad")) = 1 Then
          'Empieza la trama por Alumno
           Print #NumFileFacturas, "1";                                                  ' Localidad
           Print #NumFileFacturas, "OCP";                                                ' Transsaccion
           Print #NumFileFacturas, "SC";                                                 ' Codigo de Servicio
           Print #NumFileFacturas, "  ";                                                 ' Tipo de Cuenta
           Print #NumFileFacturas, String(8, " ");                                       ' Numero de Cuenta
           Print #NumFileFacturas, Format$(I, "0000000000000") & Format$(J, "00");         ' Valor
           Print #NumFileFacturas, CodigoCli & String(15 - Len(CodigoCli), " ");         ' Codigo del Alumno
           Print #NumFileFacturas, Codigo2 & String(20 - Len(Codigo2), " ");             ' Referencia
           Print #NumFileFacturas, "RE";                                                 ' Forma de Pago
           Print #NumFileFacturas, "USD";                                                ' Moneda
           Print #NumFileFacturas, NombreCliente & String(30 - Len(NombreCliente), " "); ' Nombre del Alumno
           Print #NumFileFacturas, String(18, " ");                                      ' Agencia de Retiro
           Print #NumFileFacturas, "0";                                                  ' NUC Ordenante
           Print #NumFileFacturas, Format$(Contador, "000000")                            ' Secuencial
        End If
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas

RutaGeneraFile2 = RutaSysBases & "\BANCO\FACTURAS\BIZBANK DEBITO DEL " _
                & Replace(MBFechaI, "/", "-") & " AL " & Replace(MBFechaF, "/", "-") & ".TXT"
RutaGeneraFile2 = UCaseStrg(RutaGeneraFile2)
NumFileFacturas = FreeFile
TipoDoc = "0"
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
Open RutaGeneraFile2 For Output As #NumFileFacturas ' Abre el archivo.
With AdoPendiente.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TxtFile = "TOTAL NOMINA DE RECAUDACION:" & vbCrLf
     Codigo3 = TrimStrg(MidStrg(NombreEmpresa, 1, 30))
     Total = 0
     TotalIngreso = 0
     Total_Factura = 0
     IE = 1
     JE = 1
     KE = 1
     Grupo_No = .fields("Grupo")
     Codigo = .fields("CodigoC")
     Do While Not .EOF
        'MsgBox .Fields("Grupo")
        'MsgBox .Fields("Grupo")
        If Grupo_No <> .fields("Grupo") Then
           Codigo4 = Format$(Total, "#,##0.00")
           Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
           TxtFile = TxtFile _
                   & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
           JE = 0
           Total = 0
           IE = IE + 1
           Grupo_No = .fields("Grupo")
           Codigo = .fields("CodigoC")
        End If
        If Codigo <> .fields("CodigoC") Then
           JE = JE + 1
           KE = KE + 1
           Codigo = .fields("CodigoC")
        End If
        FRecaudacionBancosCxC.Caption = "Por Débito: " & .fields("Grupo") & " - " & Format$(Contador / (.RecordCount * 2), "00%")
        Contador = Contador + 1
        'ProgBarra.value = Contador
        CodigoCli = .fields("CI_RUC")
        NombreCliente = Sin_Signos_Especiales(TrimStrg(MidStrg(.fields("Cliente"), 1, 30)))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & " " & Year(MBFechaI)
        Codigo2 = UCaseStrg(.fields("Grupo") & " " & Replace(MBFechaF, "/", " "))
        Total_Factura = .fields("Saldo_Pend")
        If Costo_Banco > 0 Then Total_Factura = Total_Factura + Costo_Banco
        Total = Total + Total_Factura
        TotalIngreso = TotalIngreso + Total_Factura
        I = Int(Total_Factura)
        J = (Total_Factura - Int(Total_Factura)) * 100
        If Len(.fields("Actividad")) > 1 And UCaseStrg(MidStrg(.fields("Actividad"), 1, 5)) <> "TRANS" Then
          'Empieza la trama por Alumno
           Print #NumFileFacturas, "1";                                                  ' Localidad
           Print #NumFileFacturas, "OCP";                                                ' Transsaccion
           Print #NumFileFacturas, "SC";                                                 ' Codigo de Servicio
           Print #NumFileFacturas, MidStrg(.fields("Actividad"), 1, 10);                     ' Tipo de Cuenta y Numero de Cuenta
           Print #NumFileFacturas, Format$(I, "0000000000000") & Format$(J, "00");         ' Valor
           Print #NumFileFacturas, CodigoCli & String(15 - Len(CodigoCli), " ");         ' Codigo del Alumno
           Print #NumFileFacturas, Codigo2 & String(20 - Len(Codigo2), " ");             ' Referencia
           Print #NumFileFacturas, "CU";                                                 ' Forma de Pago
           Print #NumFileFacturas, "USD";                                                ' Moneda
           Print #NumFileFacturas, NombreCliente & String(30 - Len(NombreCliente), " "); ' Nombre del Alumno
           Print #NumFileFacturas, String(18, " ");                                      ' Agencia de Retiro
           Print #NumFileFacturas, "0";                                                  ' NUC Ordenante
           Print #NumFileFacturas, Format$(Contador, "000000")                            ' Secuencial
        End If
       .MoveNext
     Loop
 End If
End With
Close #NumFileFacturas
Codigo4 = Format$(Total, "#,##0.00")
Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
TxtFile = TxtFile _
        & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf

Codigo4 = Format$(TotalIngreso, "#,##0.00")
Codigo4 = String(13 - Len(Codigo4), " ") & Format$(TotalIngreso, "#,##0.00")
TxtFile = TxtFile _
        & String(90, "-") & vbCrLf _
        & "Total Grupos: " & IE & vbTab & "Total Alumnos: " & KE & vbTab & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
FRecaudacionBancosCxC.Caption = FormCaption
''ProgBarra.value = 'ProgBarra.Max
RatonNormal
MsgBox "Fin del Proceso, el archivo se Generó en: " & vbCrLf & vbCrLf & RutaGeneraFile1 & vbCrLf & vbCrLf & RutaGeneraFile2
End Sub

Public Sub Generar_Guayaquil()
RutaGeneraFile = UCaseStrg(RutaSysBases _
               & "\BANCO\FACTURAS\RCE_" _
               & Format$(FechaSistema, "YYYYMMDD") & "_" _
               & Format$(Val(CodigoDelBanco), "0000000") _
               & "_01.TXT")
NumFileFacturas = FreeFile
TipoDoc = "0"
Contador = 0
''ProgBarra.value = 0
'ProgBarra.Min = 0
FechaTexto = BuscarFecha(MBFechaI)
'MsgBox RutaGeneraFile
TxtFile = ""
Open RutaGeneraFile For Output As #NumFileFacturas ' Abre el archivo.
With AdoPendiente.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     TxtFile = "TOTAL NOMINA DE RECAUDACION:" & vbCrLf
     'ProgBarra.Max = .RecordCount
     Codigo3 = TrimStrg(MidStrg(NombreEmpresa, 1, 30))
     Total = 0
     TotalIngreso = 0
     Total_Factura = 0
     IE = 1
     JE = 1
     KE = 1
     Grupo_No = .fields("Grupo")
     Codigo = .fields("CodigoC")
     Do While Not .EOF
        'MsgBox .Fields("Grupo")
        If Grupo_No <> .fields("Grupo") Then
           Codigo4 = Format$(Total, "#,##0.00")
           Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
           TxtFile = TxtFile _
                   & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
           JE = 0
           Total = 0
           IE = IE + 1
           Grupo_No = .fields("Grupo")
           Codigo = .fields("CodigoC")
        End If
        If Codigo <> .fields("CodigoC") Then
           JE = JE + 1
           KE = KE + 1
           Codigo = .fields("CodigoC")
        End If
        FRecaudacionBancosCxC.Caption = .fields("Grupo") & " - " & Format$(Contador / .RecordCount, "00%")
        Contador = Contador + 1
        'ProgBarra.value = Contador
        CodigoCli = CStr(Val(.fields("CI_RUC")))
        NombreCliente = Sin_Signos_Especiales(TrimStrg(MidStrg(.fields("Cliente"), 1, 40)))
        Codigo1 = TrimStrg(MidStrg(.fields("Direccion"), 1, 30))
        FechaTexto = " " & MidStrg(MesesLetras(Month(MBFechaI)), 1, 3) & " " & Year(MBFechaI)
        Codigo2 = MidStrg(UCaseStrg(MidStrg(MesesLetras(Month(MBFechaI)), 1, 3)) & " " & .fields("Grupo"), 1, 15)
        Total_Factura = .fields("Saldo_Pend")
        If Costo_Banco > 0 Then Total_Factura = Total_Factura + Costo_Banco
        Total = Total + Total_Factura
        TotalIngreso = TotalIngreso + Total_Factura
        I = Int(Total_Factura)
        J = (Total_Factura - Int(Total_Factura)) * 100
       'MsgBox Grupo_No & "(" & JE & ")" & vbCrLf & Total_Factura & vbCrLf & Total & vbCrLf & TotalIngreso
       'Empieza la trama por Alumno
       'Registro de Cobros
        Print #NumFileFacturas, "CO";                                                 ' Localidad
        Print #NumFileFacturas, Format$(Contador, "0000000");                          ' Secuencial
        Print #NumFileFacturas, CodigoCli & String(15 - Len(CodigoCli), " ");         ' Codigo del Alumno
        Print #NumFileFacturas, "USD";                                                ' Transsaccion
        Print #NumFileFacturas, Format$(I, "00000000") & Format$(J, "00");              ' Valor a Cobrar
        Print #NumFileFacturas, "REC";                                                ' Codigo de Servicio
        Print #NumFileFacturas, NombreCliente & String(40 - Len(NombreCliente), " "); ' Nombre del Alumno
        Print #NumFileFacturas, Format$(MBFechaI, "YYYYMM");                           ' Mes de Generacion
        Print #NumFileFacturas, "CU";                                                 ' Forma de Pago
        Print #NumFileFacturas, "PA";                                                 ' Moneda
        Print #NumFileFacturas, "ES";                                                 ' NUC Ordenante
        Print #NumFileFacturas, Codigo2 & String(15 - Len(Codigo2), " ")              ' Codigo del Alumno
        
       'Registro de Fecha Vencimiento
        Print #NumFileFacturas, "RC";                                                 ' Regla de Cobro
        Print #NumFileFacturas, Format$(Contador, "0000000");                          ' Secuencial
        Print #NumFileFacturas, "VM";                                                 ' Regla Cobro Fecha Vencimiento
        Print #NumFileFacturas, Format$(MBFechaI, "YYYYMMDD");                         ' Fecha Inicio de Cobro
        Print #NumFileFacturas, Format$(MBFechaF, "YYYYMMDD");                         ' Fecha Tope de Cobro
        Print #NumFileFacturas, "FI";                                                 ' Monto fijo
        Print #NumFileFacturas, String(30, "0")                                       ' Fin de la Trama
       .MoveNext
     Loop
     Codigo4 = Format$(Total, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(Total, "#,##0.00")
     TxtFile = TxtFile _
             & "Grupo: " & Grupo_No & vbTab & "Resumen de Registros por Grupo: " & JE & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
     Codigo4 = Format$(TotalIngreso, "#,##0.00")
     Codigo4 = String(13 - Len(Codigo4), " ") & Format$(TotalIngreso, "#,##0.00")
     TxtFile = TxtFile _
             & String(90, "-") & vbCrLf _
             & "Total Grupos: " & IE & vbTab & "Total Alumnos: " & KE & vbTab & vbTab & "Total a Recaudar USD" & vbTab & Codigo4 & vbCrLf
 End If
End With
Close #NumFileFacturas
''ProgBarra.value = 'ProgBarra.Max
RatonNormal
MsgBox UCaseStrg("Fin del Proceso " & vbCrLf & vbCrLf & "El Archivo se Generara en: " & vbCrLf & vbCrLf & RutaGeneraFile)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  FA.Tipo_PRN = "FM"
  Cta_Bancaria = SinEspaciosDer(DCBanco)
  Select Case Button.key
    Case "Visualizar"
         Visualizar_Archivo
    Case "Enviar"
         Enviar_Rubros
    Case "Recibir"
         Recibir_Abonos
    Case "Salir"
         Unload FRecaudacionBancosCxC
  End Select
End Sub

Private Sub TxtOrden_GotFocus()
  MarcarTexto TxtOrden
End Sub

Private Sub TxtOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOrden_LostFocus()
 If Len(TxtOrden) <= 0 Then TxtOrden = Ninguno
End Sub
