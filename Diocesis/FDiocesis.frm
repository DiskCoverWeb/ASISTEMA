VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FClientesDiocesis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "          "
   ClientHeight    =   9075
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar los Datos del Registro"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Certificado"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Crear nuevo Beneficiario"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin InetCtlsObjects.Inet URLinet 
      Left            =   6195
      Top             =   8610
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   39
      Top             =   8190
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   35
      Top             =   7875
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   34
      Top             =   7560
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   33
      Top             =   7245
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   32
      Top             =   6930
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   31
      Top             =   6615
      Width           =   7995
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   30
      Top             =   6300
      Width           =   7995
   End
   Begin VB.TextBox TxtMatricula 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5670
      MaxLength       =   10
      TabIndex        =   16
      Top             =   5145
      Width           =   1380
   End
   Begin VB.TextBox TxtPagina 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
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
      MaxLength       =   10
      TabIndex        =   14
      Top             =   5145
      Width           =   1065
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
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
      MaxLength       =   12
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   8610
      Width           =   1905
   End
   Begin VB.TextBox TxtCampo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4095
      MaxLength       =   50
      TabIndex        =   29
      Top             =   5985
      Width           =   7995
   End
   Begin VB.TextBox TxtTomo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      MaxLength       =   10
      TabIndex        =   12
      Top             =   5145
      Width           =   1170
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
      Left            =   9870
      MaxLength       =   13
      TabIndex        =   10
      ToolTipText     =   "<Ctrl+M> Codigo de Matrícula"
      Top             =   4725
      Width           =   2220
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
      TabIndex        =   8
      Top             =   4725
      Width           =   9780
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
      Height          =   3585
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   11985
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
         Height          =   330
         Left            =   9450
         MaxLength       =   13
         TabIndex        =   5
         ToolTipText     =   "<Ctrl+M> Codigo de Matrícula"
         Top             =   210
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&S"
         Height          =   330
         Left            =   11550
         TabIndex        =   37
         Top             =   210
         Width           =   330
      End
      Begin MSDataListLib.DataCombo DCCliente 
         Bindings        =   "FDiocesis.frx":0000
         DataSource      =   "AdoListCtas"
         Height          =   2910
         Left            =   105
         TabIndex        =   6
         ToolTipText     =   "Ctrl+B: Buscar datos en forma general"
         Top             =   525
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   5133
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
      Begin MSDataListLib.DataCombo DCGrupo 
         Bindings        =   "FDiocesis.frx":001A
         DataSource      =   "AdoGrupo"
         Height          =   315
         Left            =   2415
         TabIndex        =   2
         Top             =   210
         Width           =   3900
         _ExtentX        =   6879
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
      Begin MSMask.MaskEdBox MBFechaCorte 
         Height          =   330
         Left            =   8190
         TabIndex        =   4
         Top             =   210
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
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DE CERTIFICADO"
         BeginProperty Font 
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
         Top             =   210
         Width           =   2325
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de &Corte"
         BeginProperty Font 
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
         TabIndex        =   3
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc AdoBanco 
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
      Left            =   3045
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
   Begin MSAdodcLib.Adodc AdoBenef 
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
      Caption         =   "Benef"
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
      Left            =   3045
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
      Left            =   3045
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
   Begin MSMask.MaskEdBox MBFechaB 
      Height          =   330
      Left            =   4095
      TabIndex        =   28
      Top             =   5670
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
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
      Left            =   10815
      TabIndex        =   18
      Top             =   5145
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
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   7
      Left            =   105
      TabIndex        =   40
      Top             =   8190
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   6
      Left            =   105
      TabIndex        =   26
      Top             =   7875
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   5
      Left            =   105
      TabIndex        =   25
      Top             =   7560
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   4
      Left            =   105
      TabIndex        =   24
      Top             =   7245
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   3
      Left            =   105
      TabIndex        =   23
      Top             =   6930
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   2
      Left            =   105
      TabIndex        =   22
      Top             =   6615
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   1
      Left            =   105
      TabIndex        =   21
      Top             =   6300
      Width           =   4005
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA DE NACIMIENTO (dia/Mes/Año)"
      BeginProperty Font 
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
      TabIndex        =   17
      Top             =   5145
      Width           =   3690
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA BAUTIZO"
      BeginProperty Font 
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
      Top             =   5670
      Width           =   4110
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ACTA No."
      BeginProperty Font 
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
      TabIndex        =   15
      Top             =   5145
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PAGINA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   13
      Top             =   5145
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
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
      Height          =   330
      Left            =   105
      TabIndex        =   27
      Top             =   8610
      Width           =   4005
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
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
      Index           =   0
      Left            =   105
      TabIndex        =   20
      Top             =   5985
      Width           =   4005
   End
   Begin VB.Label Label42 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOMO"
      BeginProperty Font 
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
      Top             =   5145
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      X1              =   -105
      X2              =   12180
      Y1              =   5565
      Y2              =   5565
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CEDULA/NIC"
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
      Left            =   9870
      TabIndex        =   9
      Top             =   4410
      Width           =   2220
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
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   4410
      Width           =   9780
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11655
      Top             =   735
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
            Picture         =   "FDiocesis.frx":0031
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":034B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":0665
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":097F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":0C99
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":0FB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":12CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":15E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":1901
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":1C1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":1F35
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":224F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":10FE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":112FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":11615
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":1192F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":11AE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":122FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":12539
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":12853
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":12B6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":12D47
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":13061
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":1337B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":13695
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FDiocesis.frx":139AF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FClientesDiocesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Archivo_Foto As String
Dim Cliente_Ant As String
Dim NombFilePict As String
Dim AdoRegs As ADODB.Recordset
Dim Imprime As Boolean
Dim Estudiante As String
Dim SexoEst As String
Dim FechaC As String
Dim Strgs As String
Dim TipoReporte As String
Dim CodigoBenef(20) As String

Public Sub Imprimir_Certificado_Pago()
Dim tipoDeLetra As String
  Codigo = DCGrupo
  If Codigo = "" Then Codigo = Ninguno
  Mensajes = "Imprmir Certificado de " & Codigo
  Titulo = "IMPRESION"
  Bandera = False
  SetPrinters.Show 1
  If PonImpresoraDefecto(SetNombrePRN) Then
     RatonReloj
     Select Case Codigo
       Case "BAUTIZO"
            CDConLineas = ProcesarSeteos("CB")
       Case "CONFIRMACION"
            CDConLineas = ProcesarSeteos("CC")
       Case "MATRIMONIO"
            CDConLineas = ProcesarSeteos("CM")
     End Select
     tPrint.TipoImpresion = Es_Printer
     tPrint.NombreArchivo = "CERTIFICADO DE " & TxtApellidosS
     tPrint.TituloArchivo = "CERTIFICADO DE " & TxtApellidosS
     tPrint.TipoLetra = TipoCourier
     tPrint.OrientacionPagina = Orientacion_Pagina
     tPrint.PaginaA4 = True
     tPrint.EsCampoCorto = False
     tPrint.VerDocumento = True
     Set cPrint = New cImpresion
     cPrint.iniciaImpresion
     
    'Iniciamos Impresion
     tipoDeLetra = tPrint.TipoLetra
     cPrint.tipoNegrilla = True
     
     cPrint.letraTipo tipoDeLetra, SetD(1).Tamaño
     cPrint.printTexto SetD(1).PosX, SetD(1).PosY, UCase$(Empresa)
     
     Cadena = TxtCampo(3) & ", " & UCase$(FechaStrg(MBFechaB))
     cPrint.letraTipo tipoDeLetra, SetD(12).Tamaño
     cPrint.printTexto SetD(12).PosX, SetD(12).PosY, Cadena
     
     cPrint.letraTipo tipoDeLetra, SetD(13).Tamaño
     cPrint.printTexto SetD(13).PosX, SetD(13).PosY, TxtCampo(4)
    
     cPrint.letraTipo tipoDeLetra, SetD(14).Tamaño
     cPrint.printTexto SetD(14).PosX, SetD(14).PosY, TxtCampo(5)
     
     cPrint.letraTipo tipoDeLetra, SetD(16).Tamaño
     cPrint.printTexto SetD(16).PosX, SetD(16).PosY, TxtTomo
     
     cPrint.letraTipo tipoDeLetra, SetD(17).Tamaño
     cPrint.printTexto SetD(17).PosX, SetD(17).PosY, TxtPagina
     
     cPrint.letraTipo tipoDeLetra, SetD(18).Tamaño
     cPrint.printTexto SetD(18).PosX, SetD(18).PosY, TxtMatricula

     cPrint.letraTipo tipoDeLetra, SetD(5).Tamaño
     cPrint.printTexto SetD(5).PosX, SetD(5).PosY, TxtApellidosS
     
     Cadena = ULCase(NombreCiudad) & ", " & FechaStrg(FechaSistema)
     cPrint.letraTipo tipoDeLetra, SetD(19).Tamaño
     cPrint.printTexto SetD(19).PosX, SetD(19).PosY, Cadena
     
''       CodigoBenef(0) = .Fields("Padre")
''       CodigoBenef(1) = .Fields("Madre")
''       CodigoBenef(2) = .Fields("Ciudad_Nacimiento")
''       CodigoBenef(3) = .Fields("Ciudad_B_C_M")
''       CodigoBenef(4) = .Fields("Padrinos")
''       CodigoBenef(5) = .Fields("Ministro")
''       CodigoBenef(6) = .Fields("Nota_Marginal")
     
     Select Case Codigo
       Case "BAUTIZO"
            Cadena = " "
            If Len(TxtCampo(0)) > 1 And Len(TxtCampo(1)) > 1 Then Cadena = TxtCampo(0) & " Y " & TxtCampo(1)
            If Len(TxtCampo(0)) > 1 And Len(TxtCampo(1)) <= 1 Then Cadena = TxtCampo(0)
            If Len(TxtCampo(0)) <= 1 And Len(TxtCampo(1)) > 1 Then Cadena = TxtCampo(1)
            cPrint.letraTipo tipoDeLetra, SetD(7).Tamaño
            cPrint.printTexto SetD(7).PosX, SetD(7).PosY, Cadena
            
            Cadena = TxtCampo(2) & ", " & UCase$(FechaStrg(MBFechaN))
            cPrint.letraTipo tipoDeLetra, SetD(11).Tamaño
            cPrint.printTexto SetD(11).PosX, SetD(11).PosY, Cadena
            
            cPrint.letraTipo tipoDeLetra, SetD(15).Tamaño
            cPrint.printTexto SetD(15).PosX, SetD(15).PosY, TxtCampo(6)
            
            cPrint.letraTipo tipoDeLetra, SetD(20).Tamaño
            cPrint.printTexto SetD(20).PosX, SetD(20).PosY, TxtCampo(7)
       
       Case "CONFIRMACION"
            cPrint.letraTipo tipoDeLetra, SetD(7).Tamaño
            cPrint.printTexto SetD(7).PosX, SetD(7).PosY, TxtCampo(0)
            
            cPrint.letraTipo tipoDeLetra, SetD(9).Tamaño
            cPrint.printTexto SetD(9).PosX, SetD(9).PosY, TxtCampo(1)
                               
       Case "MATRIMONIO"
            cPrint.letraTipo tipoDeLetra, SetD(9).Tamaño
            cPrint.printTexto SetD(9).PosX, SetD(9).PosY, TxtCampo(1)
       
     End Select
    'fin del documento
     cPrint.finalizaImpresion
     RatonNormal
 End If
End Sub

Public Sub GrabarCliente()
  Si_No = False
  T = Normal
  Mensajes = "Esta seguro de Grabar datos de:" & vbCrLf _
           & TxtApellidosS & vbCrLf
  Titulo = "Pregunta de grabación"
  If BoxMensaje Then
     Codigo = DCGrupo
     If Codigo = "" Then Codigo = Ninguno
     sSQL = "SELECT * " _
          & "FROM Trans_Parroquias " _
          & "WHERE Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Tipo_Certificado = '" & Codigo & "' " _
          & "AND Beneficiario = '" & TxtApellidosS & "' "
     Select_Adodc AdoBenef, sSQL
     With AdoBenef.Recordset
      If .RecordCount > 0 Then
         .Fields("Tomo") = TxtTomo
         .Fields("Pagina") = TxtPagina
         .Fields("Numero") = TxtMatricula
         .Fields("Fecha_Nacimiento") = MBFechaN
         .Fields("Fecha_B_C_M") = MBFechaB
         .Fields("Padre") = TxtCampo(0)
         .Fields("Madre") = TxtCampo(1)
         .Fields("Ciudad_Nacimiento") = TxtCampo(2)
         .Fields("Ciudad_B_C_M") = TxtCampo(3)
         .Fields("Padrinos") = TxtCampo(4)
         .Fields("Ministro") = TxtCampo(5)
         .Fields("Nota_Marginal") = TxtCampo(6)
         .Fields("Certificado_Valido") = TxtCampo(7)
         .Update
      Else
          SetAdoAddNew "Trans_Parroquias"
          SetAdoFields "T", Normal
          SetAdoFields "Tipo_Certificado", Codigo
          SetAdoFields "Fecha", FechaSistema
          SetAdoFields "Cedula", Ninguno
          SetAdoFields "Cedula_P", Ninguno
          SetAdoFields "Cedula_M", Ninguno
          SetAdoFields "Beneficiario", TxtApellidosS
          SetAdoFields "Tomo", Val(TxtTomo)
          SetAdoFields "Pagina", Val(TxtPagina)
          SetAdoFields "Numero", Val(TxtMatricula)
          SetAdoFields "Fecha_Nacimiento", MBFechaN
          SetAdoFields "Fecha_B_C_M", MBFechaB
          SetAdoFields "Padre", TxtCampo(0)
          SetAdoFields "Madre", TxtCampo(1)
          SetAdoFields "Ciudad_Nacimiento", TxtCampo(2)
          SetAdoFields "Ciudad_B_C_M", TxtCampo(3)
          SetAdoFields "Padrinos", TxtCampo(4)
          SetAdoFields "Ministro", TxtCampo(5)
          SetAdoFields "Nota_Marginal", TxtCampo(6)
          SetAdoFields "Certificado_Valido", TxtCampo(7)
          SetAdoUpdate
          ListarDiocesis
      End If
     End With
     RatonNormal
     MsgBox "Proceso Exitoso"
  End If
End Sub

Public Sub ListarDiocesis()
Dim TextosCliente As String
  RatonReloj
  Codigo = DCGrupo
  If Codigo = "" Then Codigo = "BAUTIZO"
  sSQL = "SELECT Beneficiario, ID " _
       & "FROM Trans_Parroquias " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Certificado = '" & Codigo & "' " _
       & "ORDER BY Beneficiario "
  SelectDB_Combo DCCliente, AdoListCtas, sSQL, "Beneficiario"
  Frame1.Caption = " NOMBRES DE " & Codigo & " " & Space(30) & "Total Clientes: " & Format$(AdoListCtas.Recordset.RecordCount, "000000")
  RatonNormal
 'DCCliente.SetFocus
End Sub

Public Sub ListarCuenta(TextoBusqueda As String)
  TxtCodigo = Ninguno
  TxtDescuento = "0.00"
  Codigo = DCGrupo
  If Codigo = "" Then Codigo = Ninguno
  sSQL = "SELECT * " _
       & "FROM Trans_Parroquias " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Certificado = '" & Codigo & "' " _
       & "AND Beneficiario = '" & TextoBusqueda & "' "
  Select_Adodc AdoBenef, sSQL
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       TxtCodigo = .Fields("ID")
       TxtApellidosS = .Fields("Beneficiario")
       TxtTomo = .Fields("Tomo")
       TxtPagina = .Fields("Pagina")
       TxtMatricula = .Fields("Numero")
       MBFechaCorte = .Fields("Fecha")
       MBFechaN = .Fields("Fecha_Nacimiento")
       MBFechaB = .Fields("Fecha_B_C_M")
       CodigoBenef(0) = .Fields("Padre")
       CodigoBenef(1) = .Fields("Madre")
       CodigoBenef(2) = .Fields("Ciudad_Nacimiento")
       CodigoBenef(3) = .Fields("Ciudad_B_C_M")
       CodigoBenef(4) = .Fields("Padrinos")
       CodigoBenef(5) = .Fields("Ministro")
       CodigoBenef(6) = .Fields("Nota_Marginal")
       CodigoBenef(7) = .Fields("Certificado_Valido")
   Else
      MsgBox "No Existe"
   End If
  End With
  Select Case Codigo
    Case "BAUTIZO"
         Label15.Visible = True
         MBFechaN.Visible = True
         Label25(0).Visible = True
         Label25(2).Visible = True
         Label25(6).Visible = True
         Label25(7).Visible = True
         TxtCampo(0).Visible = True
         TxtCampo(2).Visible = True
         TxtCampo(6).Visible = True
         TxtCampo(7).Visible = True
         Label25(0).Caption = " NOMBRE DEL PADRE"
         Label25(1).Caption = " NOMBRE DE LA MADRE"
         Label25(2).Caption = " CIUDAD DE NACIMIENTO"
         Label25(3).Caption = " CIUDAD DE BAUTIZO"
         Label25(4).Caption = " PADRINO(S)"
         Label25(5).Caption = " MINISTRO"
         Label25(6).Caption = " NOTA MARGINAL"
         Label25(7).Caption = " CERTIFICADO VALIDO PARA"
    Case "CONFIRMACION"
         Label15.Visible = False
         MBFechaN.Visible = False
         Label25(0).Visible = True
         Label25(2).Visible = False
         Label25(6).Visible = False
         Label25(7).Visible = False
         TxtCampo(0).Visible = True
         TxtCampo(2).Visible = False
         TxtCampo(6).Visible = False
         TxtCampo(7).Visible = False
         Label25(0).Caption = " NOMBRE DEL PADRE"
         Label25(1).Caption = " NOMBRE DE LA MADRE"
         Label25(3).Caption = " CIUDAD DE CONFIRMACION"
         Label25(4).Caption = " PADRINOS"
         Label25(5).Caption = " MINISTRO"
         CodigoBenef(6) = Ninguno
    Case "MATRIMONIO"
         Label15.Visible = False
         MBFechaN.Visible = False
         Label25(0).Visible = False
         Label25(2).Visible = False
         Label25(6).Visible = False
         TxtCampo(0).Visible = False
         TxtCampo(2).Visible = False
         TxtCampo(6).Visible = False
         Label25(1).Caption = " NOMBRE DE LA ESPOSA"
         Label25(3).Caption = " CIUDAD DE MATRIMONIO"
         Label25(4).Caption = " PADRINOS"
         Label25(5).Caption = " MINISTRO"
         CodigoBenef(0) = Ninguno
         CodigoBenef(6) = Ninguno
  End Select
  Label14.Caption = " FECHA DE " & Codigo & " (Dia/Mes/Año)"
  For I = 0 To 7
      TxtCampo(I) = CodigoBenef(I)
  Next I
End Sub

Private Sub Command1_Click()
  Unload FClientesDiocesis
End Sub

Private Sub DCCliente_DblClick(Area As Integer)
  SiguienteControl
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
Dim NombreTabla(20) As String
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyB Then
     ListarDiocesis
     MsgBox "Busque el dato"
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
  ListarDiocesis
End Sub

Private Sub Form_Activate()
  FechaComp = FechaSistema
  MBFechaCorte = "30/06/2017"
  For I = 0 To UBound(CodigoBenef) - 1
      CodigoBenef(I) = Ninguno
  Next I
  
  sSQL = "SELECT Tipo_Certificado " _
       & "FROM Trans_Parroquias " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "GROUP BY Tipo_Certificado " _
       & "ORDER BY Tipo_Certificado "
  SelectDB_Combo DCGrupo, AdoGrupo, sSQL, "Tipo_Certificado"
  RatonNormal
  DCGrupo.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FClientesDiocesis
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoBenef
   ConectarAdodc AdoGrupo
   ConectarAdodc AdoCuentas
   ConectarAdodc AdoListCtas
   ConectarAdodc AdoBanco
   ConectarAdodc AdoCreditos
   ConectarAdodc AdoDireccion
   FClientesDiocesis.Caption = "GENERACIONES DE BAUTIZOS-CONFIRMACIONES-MATRIMONIOS"
End Sub

Private Sub MBFechaB_GotFocus()
  MarcarTexto MBFechaB
End Sub

Private Sub MBFechaB_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaB_LostFocus()
  FechaValida MBFechaB
End Sub

Private Sub MBFechaN_GotFocus()
  MarcarTexto MBFechaN
End Sub

Private Sub MBFechaN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaN_LostFocus()
  FechaValida MBFechaN
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
 
'MsgBox Button.key
 Select Case Button.key
   Case "Salir"
        RatonNormal
        Unload FClientesDiocesis
   Case "Grabar"
        GrabarCliente
   Case "Imprimir"
        Imprimir_Certificado_Pago
   Case "Eliminar"
        Mensajes = "Esta seguro de Eliminar datos de:" & vbCrLf _
                 & TxtApellidosS & vbCrLf
        Titulo = "PREGUNTA DE ELIMINACION"
        If BoxMensaje Then
           Mensajes = "ADVERTENCIA: " & vbCrLf _
                    & "ESTA SEGURO DE QUERER ELIMINAR LOS DATOS DE:" & vbCrLf & TxtApellidosS & vbCrLf _
                    & "NO PODRA VOLVER A RECUPERAR ESTOS DATOS" & vbCrLf
           If BoxMensaje Then
              Codigo = DCGrupo
              If Codigo = "" Then Codigo = Ninguno
              sSQL = "DELETE * " _
                   & "FROM Trans_Parroquias " _
                   & "WHERE Periodo = '" & Periodo_Contable & "' " _
                   & "AND Item = '" & NumEmpresa & "' " _
                   & "AND Tipo_Certificado = '" & Codigo & "' " _
                   & "AND Beneficiario = '" & TxtApellidosS & "' "
              Ejecutar_SQL_SP sSQL
              ListarDiocesis
           End If
        End If
   Case "Nuevo"
        TxtApellidosS = ""
        TxtTomo = "0"
        TxtPagina = "0"
        TxtMatricula = "0"
        MBFechaCorte = FechaSistema
        MBFechaN = "00/00/0000"
        MBFechaB = "00/00/0000"
        For I = 0 To 6
            TxtCampo(I) = ""
        Next I
        MsgBox "Ingrese los datos del Beneficiario:"
        TxtApellidosS.SetFocus
 End Select
 RatonNormal
End Sub


Private Sub TxtCampo_GotFocus(Index As Integer)
   MarcarTexto TxtCampo(Index)
End Sub

Private Sub TxtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCampo_LostFocus(Index As Integer)
  TextoValido TxtCampo(Index), , True
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
'''  TextoValido TxtApellidosS, , True
'''  With AdoListCtas.Recordset
'''   If .RecordCount > 0 And TxtApellidosS.Text <> Ninguno Then
'''       RatonReloj
'''      .MoveFirst
'''      .Find ("Cliente Like '" & TxtApellidosS & "' ")
'''       RatonNormal
'''       If Not .EOF Then
'''          MsgBox "El Cliente " & TxtApellidosS _
'''               & ", ya existe, está asignado a " & vbCrLf & vbCrLf _
'''               & .Fields("Cliente") & vbCrLf & vbCrLf _
'''               & "Codigo: " & .Fields("CI_RUC")
'''          DCCliente.SetFocus
'''       End If
'''   End If
'''  End With
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
          DigVerif = Digito_Verificador(TxtCI_RUC.Text)
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

Private Sub TxtCodigo_GotFocus()
  MarcarTexto TxtCodigo
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDescuento_GotFocus()
   MarcarTexto TxtDescuento
End Sub

Private Sub TxtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaCorte_GotFocus()
   MarcarTexto MBFechaCorte
End Sub

Private Sub MBFechaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBFechaCorte_LostFocus()
  FechaValida MBFechaCorte
End Sub

Private Sub TxtMatricula_GotFocus()
   MarcarTexto TxtMatricula
End Sub

Private Sub TxtMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtPagina_GotFocus()
  MarcarTexto TxtPagina
End Sub

Private Sub TxtPagina_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtTomo_GotFocus()
  MarcarTexto TxtTomo
End Sub

Private Sub TxtTomo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
