VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRetIVAF 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobante de Retención en la Fuente e IVA"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   Icon            =   "FrmRetIVAF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9173.079
   ScaleMode       =   0  'User
   ScaleWidth      =   11983.23
   Begin VB.TextBox TextIVA 
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
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   2730
      MaxLength       =   15
      TabIndex        =   52
      Text            =   "0.00"
      Top             =   5145
      Width           =   1380
   End
   Begin VB.TextBox TxBaseImpo 
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
      Left            =   8295
      MaxLength       =   15
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   1050
      Width           =   1380
   End
   Begin VB.TextBox TextNumCR 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   9135
      MaxLength       =   10
      TabIndex        =   35
      Text            =   "0000000"
      Top             =   1890
      Width           =   1485
   End
   Begin VB.ComboBox CCtaIVA70 
      BackColor       =   &H00FF8080&
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
      Height          =   315
      Left            =   1575
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1050
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.TextBox TxBaseImpoF 
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
      Left            =   6825
      MaxLength       =   15
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   1050
      Width           =   1485
   End
   Begin VB.ComboBox CCtaIVA30 
      BackColor       =   &H00FF8080&
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
      Height          =   315
      Left            =   1575
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   735
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.ComboBox CCtaRet 
      BackColor       =   &H00FF8080&
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
      Height          =   315
      Left            =   1575
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   420
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   645
      Left            =   105
      TabIndex        =   78
      Top             =   6300
      Width           =   8625
      Begin VB.Label LblProv 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor:"
         BeginProperty Font 
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
         TabIndex        =   80
         Top             =   210
         Width           =   7050
      End
      Begin VB.Label LProCli 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor:"
         BeginProperty Font 
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
         TabIndex        =   79
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.TextBox TextSec 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   27
      Text            =   "0000000"
      Top             =   1890
      Width           =   960
   End
   Begin VB.TextBox TextAuto 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5985
      MaxLength       =   10
      TabIndex        =   31
      Text            =   "0000000000"
      Top             =   1890
      Width           =   1275
   End
   Begin VB.TextBox TextSerie 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5145
      MaxLength       =   10
      TabIndex        =   29
      Text            =   "000000"
      Top             =   1890
      Width           =   855
   End
   Begin VB.TextBox TextAdu 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   7245
      MaxLength       =   16
      TabIndex        =   33
      Text            =   "0000000000000000"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ComboBox ComboCT 
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
      TabIndex        =   44
      Text            =   "00      No Aplica"
      Top             =   3465
      Width           =   10515
   End
   Begin VB.TextBox Txtbasecero 
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
      Left            =   4095
      MaxLength       =   15
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   5145
      Width           =   1380
   End
   Begin VB.TextBox TxtICE 
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
      Left            =   5460
      MaxLength       =   15
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   5145
      Width           =   1275
   End
   Begin VB.ComboBox CPorcIVA 
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
      Height          =   315
      Left            =   105
      TabIndex        =   48
      Top             =   5145
      Width           =   1380
   End
   Begin VB.ComboBox CPorcRet 
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
      Height          =   315
      Left            =   1470
      TabIndex        =   50
      Top             =   5145
      Width           =   1275
   End
   Begin VB.ComboBox ComboICE 
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
      Height          =   315
      Left            =   6720
      TabIndex        =   58
      Top             =   5145
      Width           =   2640
   End
   Begin VB.TextBox TextValorICE 
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
      Left            =   9345
      MaxLength       =   15
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   5145
      Width           =   1275
   End
   Begin VB.TextBox Text3 
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
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   4095
      MaxLength       =   15
      TabIndex        =   69
      Text            =   "0000000000000"
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00FF8080&
      Caption         =   "&Salir"
      DisabledPicture =   "FrmRetIVAF.frx":030A
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
      Left            =   9765
      MouseIcon       =   "FrmRetIVAF.frx":0D54
      Picture         =   "FrmRetIVAF.frx":179E
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00FF8080&
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
      Left            =   8820
      Picture         =   "FrmRetIVAF.frx":1BE0
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   6300
      Width           =   855
   End
   Begin VB.TextBox TValRetTransf 
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
      Left            =   2730
      MaxLength       =   10
      TabIndex        =   66
      Text            =   "0.00"
      Top             =   5880
      Width           =   1380
   End
   Begin VB.TextBox TValRetServ 
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
      Left            =   6720
      MaxLength       =   15
      TabIndex        =   73
      Text            =   "0.00"
      Top             =   5880
      Width           =   1275
   End
   Begin VB.TextBox TxtValorIVA 
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
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      MaxLength       =   15
      TabIndex        =   62
      Text            =   "0.00"
      Top             =   5880
      Width           =   1380
   End
   Begin VB.ComboBox CPorRetTransf 
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
      Height          =   315
      Left            =   1470
      TabIndex        =   64
      Top             =   5880
      Width           =   1275
   End
   Begin VB.TextBox TxtVRet 
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
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   9345
      MaxLength       =   15
      TabIndex        =   77
      Text            =   "0.00"
      Top             =   5880
      Width           =   1275
   End
   Begin VB.TextBox TBaseServ 
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
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   4095
      MaxLength       =   15
      TabIndex        =   68
      Text            =   "0.00"
      Top             =   5880
      Width           =   1275
   End
   Begin VB.ComboBox CPorRetServ 
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
      Height          =   315
      Left            =   5355
      TabIndex        =   71
      Top             =   5880
      Width           =   1380
   End
   Begin VB.Frame FrmDev 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dev."
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
      Left            =   9765
      TabIndex        =   13
      Top             =   0
      Width           =   855
      Begin VB.OptionButton OpcD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Si"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   210
         Width           =   540
      End
      Begin VB.OptionButton OpcD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   525
         Value           =   -1  'True
         Width           =   540
      End
   End
   Begin VB.Frame FrmPres 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Presuntivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6825
      TabIndex        =   10
      Top             =   0
      Width           =   2850
      Begin VB.OptionButton OptP 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Si"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OptP 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No"
         Height          =   225
         Index           =   1
         Left            =   1470
         TabIndex        =   12
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Compras"
      BeginProperty Font 
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
      TabIndex        =   0
      Top             =   0
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ventas"
      BeginProperty Font 
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
      Left            =   1365
      TabIndex        =   1
      Top             =   0
      Width           =   1065
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exportaciones"
      BeginProperty Font 
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
      Left            =   4305
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.OptionButton Opctrans 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Importaciones"
      BeginProperty Font 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoRetenido 
      Height          =   330
      Left            =   2100
      Top             =   6615
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
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc AdoComprobante 
      Height          =   330
      Left            =   5880
      Top             =   6615
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
      Caption         =   "NumEg"
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
   Begin MSAdodcLib.Adodc AdoConcepto 
      Height          =   330
      Left            =   105
      Top             =   6615
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
   Begin MSAdodcLib.Adodc AdoRetencion 
      Height          =   330
      Left            =   3990
      Top             =   6615
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
   Begin MSAdodcLib.Adodc AdoSCT 
      Height          =   330
      Left            =   5880
      Top             =   6300
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
      Caption         =   "Query1"
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
   Begin MSDataListLib.DataCombo DCComprobante 
      Bindings        =   "FrmRetIVAF.frx":2022
      DataSource      =   "AdoComprobante"
      Height          =   315
      Left            =   105
      TabIndex        =   37
      Top             =   2520
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCConcepto 
      Bindings        =   "FrmRetIVAF.frx":203F
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   105
      TabIndex        =   46
      Top             =   4410
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCCompMod 
      Bindings        =   "FrmRetIVAF.frx":2059
      DataSource      =   "AdoComprobante"
      Height          =   315
      Left            =   1575
      TabIndex        =   42
      Top             =   3150
      Visible         =   0   'False
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCSCT 
      Bindings        =   "FrmRetIVAF.frx":2076
      DataSource      =   "AdoSCT"
      Height          =   315
      Left            =   105
      TabIndex        =   45
      Top             =   3780
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DCNumMod 
      Bindings        =   "FrmRetIVAF.frx":208B
      DataSource      =   "AdoConcepto"
      Height          =   315
      Left            =   105
      TabIndex        =   39
      Top             =   3150
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
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
   Begin MSMask.MaskEdBox MBoxFechaE 
      Height          =   330
      Left            =   105
      TabIndex        =   21
      Top             =   1890
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
   Begin MSMask.MaskEdBox MBoxFechaR 
      Height          =   330
      Left            =   1470
      TabIndex        =   23
      Top             =   1890
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
   Begin MSMask.MaskEdBox MBoxFechaC 
      Height          =   330
      Left            =   2835
      TabIndex        =   25
      Top             =   1890
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor de IVA"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2730
      TabIndex        =   51
      Top             =   4725
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible Gravada"
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
      Left            =   8295
      TabIndex        =   18
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retencion No"
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
      Left            =   9135
      TabIndex        =   34
      Top             =   1470
      Width           =   1485
   End
   Begin MSForms.CheckBox CheckBox3 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1050
      Width           =   1485
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2619;582"
      Value           =   "0"
      Caption         =   "Ret. Servicio"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible Retencion"
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
      Left            =   6825
      TabIndex        =   16
      Top             =   630
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emision Comprobante:"
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
      TabIndex        =   20
      Top             =   1470
      Width           =   1380
   End
   Begin MSForms.CheckBox CheckBox2 
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   1485
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2619;582"
      Value           =   "0"
      Caption         =   "Ret. Bienes"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   420
      Width           =   1485
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2619;582"
      Value           =   "0"
      Caption         =   "Ret. Fuente"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Serie Factura"
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
      Left            =   5145
      TabIndex        =   28
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Autorizacion Factura"
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
      Left            =   5985
      TabIndex        =   30
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factura No"
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
      Left            =   4200
      TabIndex        =   26
      Top             =   1470
      Width           =   960
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registro Comprobante:"
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
      Left            =   1470
      TabIndex        =   22
      Top             =   1470
      Width           =   1380
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Caducidad"
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
      Left            =   2835
      TabIndex        =   24
      Top             =   1470
      Width           =   1380
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Refrendo"
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
      TabIndex        =   32
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Crédito Tributario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   40
      Top             =   2835
      Width           =   7575
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero:"
      BeginProperty Font 
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
      TabIndex        =   38
      Top             =   2835
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comprobante de Compra:"
      BeginProperty Font 
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
      TabIndex        =   36
      Top             =   2205
      Width           =   10515
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto de Retención en la Fuente:"
      BeginProperty Font 
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
      TabIndex        =   43
      Top             =   4095
      Width           =   10515
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C Modificado :"
      BeginProperty Font 
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
      TabIndex        =   41
      Top             =   2835
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje de IVA%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   47
      Top             =   4725
      Width           =   1380
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Base Imponible     Tarifa CERO (0%)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4095
      TabIndex        =   53
      Top             =   4725
      Width           =   1380
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto de ICE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5460
      TabIndex        =   55
      Top             =   4725
      Width           =   1275
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porc. % Retencion en la Fuente"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1470
      TabIndex        =   49
      Top             =   4725
      Width           =   1275
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentajes de ICE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   57
      Top             =   4725
      Width           =   2640
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor de ICE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9345
      TabIndex        =   59
      Top             =   4725
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto IVA Tranf. de Bienes:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   61
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto Retenido en Tranf. Bienes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2730
      TabIndex        =   65
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porc % IVA   Prest. Servicios"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5355
      TabIndex        =   70
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto IVA en Prest. Servicios"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4095
      TabIndex        =   67
      Top             =   5460
      Width           =   1275
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje IVA Transf.  Bienes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1470
      TabIndex        =   63
      Top             =   5460
      Width           =   1275
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monto Retenido Prest. Servicios"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   72
      Top             =   5460
      Width           =   1275
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IVA PAGADO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7980
      TabIndex        =   74
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Label LApagar 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   7980
      TabIndex        =   75
      Top             =   5880
      Width           =   1380
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RETENIDO EN LA FUENTE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9345
      TabIndex        =   76
      Top             =   5460
      Width           =   1275
   End
End
Attribute VB_Name = "FRetIVAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ErrorCT()
   MsgBox "El credito tributario elejido no corresponde a la transaccion"
   ''DCSCT.Visible = False
   ComboCT.Visible = True
   ComboCT.SetFocus
End Function

Private Sub CheckBox1_Click()
If CheckBox1.Value Then CCtaRet.Visible = True Else CCtaRet.Visible = False
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value Then CCtaIVA30.Visible = True Else CCtaIVA30.Visible = False
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value Then CCtaIVA70.Visible = True Else CCtaIVA70.Visible = False
End Sub

Private Sub CmdGrabar_Click()
  sSQL = "DELETE * " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND Cta = '" & Ninguno & "' " _
       & "AND T_No = " & Trans_No & " "
  ConectarAdoExecute sSQL
 'Retencion en la Fuente
  If CheckBox1.Value Then
     Cta = SinEspaciosIzq(CCtaRet)
     sSQL = "DELETE * " _
          & "FROM Asiento_R " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND Cta = '" & Cta & "' " _
          & "AND TD = '" & SinEspaciosIzq(DCConcepto) & "' " _
          & "AND T_No = " & Trans_No & " "
     ConectarAdoExecute sSQL
  End If
 'Retencion IVA Bienes
  If CheckBox2.Value Then
     Cta = SinEspaciosIzq(CCtaIVA30)
     sSQL = "DELETE * " _
          & "FROM Asiento_R " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND Cta = '" & Cta & "' " _
          & "AND T_No = " & Trans_No & " "
     ConectarAdoExecute sSQL
  End If
 'Retencion IVA Servicios
  If CheckBox3.Value Then
     Cta1 = SinEspaciosIzq(CCtaIVA70)
     sSQL = "DELETE * " _
          & "FROM Asiento_R " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND Cta1 = '" & Cta1 & "' " _
          & "AND T_No = " & Trans_No & " "
     ConectarAdoExecute sSQL
  End If
 'Grabar Retencion del IVA
     SetAdoAddNew "Asiento_R"
     SetAdoFields "CodigoTR", "IV"
     SetAdoFields "Retencion_No", TextNumCR
     Select Case Topc
       Case "1"
            SetAdoFields "TT", "C"
       Case "2"
            SetAdoFields "TT", "V"
            MBoxFechaC = UltimoDiaMes(MBoxFechaR)
            If Val(TextSec) > 1 And (MsgBox("Facturacion Individual?", vbYesNo) = vbYes) Then
               SetAdoFields "NumeroR", 1
            Else
               SetAdoFields "NumeroR", Val(TextSec)
            End If
       Case "3"
            SetAdoFields "TT", "I"
       Case "4"
            SetAdoFields "TT", "E"
     End Select
     SetAdoFields "Fecha", MBoxFechaR
     SetAdoFields "Valor_Fact", TxBaseImpo
     SetAdoFields "Porc", PorcIva / 100
     SetAdoFields "Valor_IVA", CCur(Val(TextIVA))  'Valor del iva
     SetAdoFields "Valor_Ret", CCur(Val(TValRetTransf)) + CCur(Val(TValRetServ))
     SetAdoFields "BImpotcero", Txtbasecero
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Codigo", CodigoCliente
     SetAdoFields "TD", SinEspaciosIzq(DCComprobante)
     SetAdoFields "FechaE", MBoxFechaE
     SetAdoFields "Autorizacion", TextAuto
     SetAdoFields "Serie", TextSerie
     SetAdoFields "Secuencial", TextSec
     SetAdoFields "IdenCT", SinEspaciosIzq(ComboCT)
     SetAdoFields "PorRetIVA1", CodPorRetTrans
     SetAdoFields "PorRetIVA2", CodPorRetSer
     SetAdoFields "Aduana", TextAdu
     If OpcD(0) Then SetAdoFields "Dev", "S" Else SetAdoFields "Dev", "N"
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "SerieEx", Text3.Text
     If Label35.Visible = True Then
        SetAdoFields "Cambio", SinEspaciosIzq(DCCompMod) & DCNumMod
     Else
        SetAdoFields "Cambio", "0"
     End If
     If OptP(0) Then
        SetAdoFields "ConvInt", "S"
     Else
        SetAdoFields "ConvInt", "N"
     End If
     SetAdoFields "CodPorc", CodPorcIva
     SetAdoFields "T_No", Trans_No
     SetAdoFields "FechaC", MBoxFechaC
     SetAdoFields "FechaA", "01/01/1999"
     SetAdoFields "CPorcICE", SinEspaciosIzq(ComboICE)
     SetAdoFields "RetICE", TextValorICE
     SetAdoFields "SCT", SinEspaciosIzq(DCSCT)
     SetAdoFields "MontoIVA1", TxtValorIVA
     SetAdoFields "MontoIVA2", TBaseServ
     SetAdoFields "MontoRetIVA2", TValRetServ
     SetAdoFields "MontoRetIVA1", TValRetTransf
     If ComboICE.ListIndex < 0 Then Cadena2 = "00" Else Cadena2 = Format(ComboICE.ListIndex, "00")
     SetAdoFields "CPorcICE", Cadena2
     If CheckBox2.Value Then SetAdoFields "Cta", SinEspaciosIzq(CCtaIVA30)
     If CheckBox3.Value Then SetAdoFields "Cta1", SinEspaciosIzq(CCtaIVA70)
     SetAdoUpdate
  'End If
 'Grabar Retencion en la Fuente
  If CheckBox1.Value Then
     SetAdoAddNew "Asiento_R"
     Select Case Topc
       Case "1"
            SetAdoFields "CodigoTR", "RF"
            SetAdoFields "TT", "C"
       Case "2"
            SetAdoFields "CodigoTR", "RV"
            SetAdoFields "TT", "V"
       Case "3"
            SetAdoFields "CodigoTR", "RI"
            SetAdoFields "TT", "I"
       Case "4"
            SetAdoFields "CodigoTR", "RE"
            SetAdoFields "TT", "E"
     End Select
     SetAdoFields "Fecha", MBoxFechaR
     SetAdoFields "Retencion_No", TextNumCR
     BaseImpon = 0
     BaseImpon = BaseImpon + Val(CCur(TxBaseImpoF))
    '' BaseImpon = BaseImpon + Txtbasecero
    'MsgBox BaseImpon
     SetAdoFields "Valor_Fact", BaseImpon
     SetAdoFields "Porc", PorRet / 100
     SetAdoFields "Valor_Ret", CCur(Val(TxtVRet))
     SetAdoFields "BImpotcero", 0
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Codigo", CodigoCliente
     SetAdoFields "TD", SinEspaciosIzq(DCConcepto)
     SetAdoFields "FechaE", MBoxFechaE
     SetAdoFields "Autorizacion", AutorizaRet
     SetAdoFields "Serie", SerieRet
     SetAdoFields "Secuencial", TextNumCR
     SetAdoFields "IdenCT", "00"
     SetAdoFields "PorRetIVA1", "0"
     SetAdoFields "PorRetIVA2", "0"
     SetAdoFields "Aduana", "0000000000000000"
     SetAdoFields "Dev", "N"
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "SerieEx", "0"
     SetAdoFields "Cambio", "0"
     SetAdoFields "ConvInt", "N"
     SetAdoFields "CodPorc", CodPorc
     SetAdoFields "T_No", Trans_No
     SetAdoFields "FechaC", "01/01/1999"
     SetAdoFields "FechaA", "01/01/1999"
     SetAdoFields "CPorcICE", "00"
     SetAdoFields "RetICE", 0
     SetAdoFields "SCT", "00"
     SetAdoFields "Cta", SinEspaciosIzq(CCtaRet)
     SetAdoUpdate
  End If
  Unload FRetIVAF
End Sub

Private Sub CmdSalir_Click()
  Unload FRetIVAF
End Sub

Private Sub ComboCT_GotFocus()
  MarcarTexto ComboCT
  If Topc = 1 Then ComboCT.Text = ComboCT.List(1)
End Sub

Private Sub ComboCT_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub ComboCT_LostFocus()
  Select Case Topc
      Case 1
          If SinEspaciosIzq(ComboCT) = "00" Then ErrorCT
      Case 2
          If SinEspaciosIzq(ComboCT) <> "00" Then ErrorCT
          ComboCT.Text = ComboCT.List(0)
   End Select
  Select Case SinEspaciosIzq(ComboCT)
  Case "02", "04", "05", "07"
    DCSCT.Text = "00    Ninguno"
    DCSCT.Enabled = False
  Case Else
    DCSCT.Enabled = True
  End Select
End Sub

Private Sub ComboICE_GotFocus()
  Cadena = BuscarFecha("31/07/2004") 'MsgBox Periodo & Cadena
  If Periodo > Cadena Or BuscarFecha(MBoxFechaE) > Cadena Then
     ComboICE.AddItem "98.00 % Cigarrillos Rubios"
     ComboICE.AddItem "32.00 % Alcohol y productos alcoholicos distintos a la cerveza"
  Else
     ComboICE.AddItem "77.25 % Cigarrillos Rubios"
     ComboICE.AddItem "26.78 % Alcohol y productos alcoholicos distintos a la cerveza"
  End If
  ComboICE.Text = ComboICE.List(0)
End Sub

Private Sub ComboICE_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub ComboICE_LostFocus()
Select Case Val(SinEspaciosIzq(ComboICE))
  Case 0: CodICE = "00"
  Case 77.25: CodICE = "01"
  Case 18.54: CodICE = "02"
  Case 30.9: CodICE = "03"
  Case 10.3
    If Len(CodICE) > 30 Then
       CodICE = "07"
    Else
       CodICE = "04"
    End If
  Case 26.78: CodICE = "05"
  Case 5.15: CodICE = "06"
  Case 15: CodICE = "08"
  Case 98: CodICE = "09"
  Case 32: CodICE = "10"
End Select
End Sub

Private Sub CPorcIVA_GotFocus()
  CPorcIVA.Text = CPorcIVA.List(0)
  MarcarTexto CPorcIVA
End Sub

Private Sub CPorcIVA_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CPorcIVA_LostFocus()
Porc = SinEspaciosIzq(CPorcIVA.Text)
Select Case Porc
 Case 10: CodPorcIva = 1
 Case 12: CodPorcIva = 2
 Case 14: CodPorcIva = 3
 Case Else: CodPorcIva = 0
End Select
If Porc <> 0 Then
PorcIva = Porc
Else
  PorcIva = 0
End If
PorcIva = Val(PorcIva)
IvaTotal = 0
'MsgBox BaseImpo
IvaTotal = Round(BaseImpo * PorcIva / 100, 2)
End Sub

Private Sub CPorcRet_GotFocus()
  CPorcRet.Text = CPorcRet.List(0)
  MarcarTexto CPorcRet
End Sub

Private Sub CPorcRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CPorcRet_LostFocus()
    PorRet = SinEspaciosIzq(CPorcRet & " ")
    Select Case PorRet
      Case 1:   CodPorc = 1
      Case 5:   CodPorc = 2
      Case 8:   CodPorc = 3
      Case Else:   CodPorc = 0
    End Select
End Sub

Private Sub CPorRetServ_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CPorRetServ_LostFocus()
  If Val(CPorRetServ) = 30 Then MsgBox "El porcentaje 30% no corresponde al tipo de transaccion"
  If CPorRetServ.ListIndex < 0 Then
     CodPorRetSer = 0
  Else
     CodPorRetSer = CPorRetServ.ListIndex
  End If
  PorRetSer = Val(Mid(CPorRetServ.Text, 1, 5))
  IvaServ = 0
  IvaServ = BaseServ * PorRetSer / 100
End Sub

Private Sub CPorRetTransf_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CPorRetTransf_GotFocus()
  CPorRetTransf.Clear
  CPorRetTransf.AddItem "0 %"
  CPorRetTransf.AddItem "30 %"
  CPorRetTransf.AddItem "70 %"
  CPorRetTransf.AddItem "100 %"
  CPorRetTransf.Text = CPorRetTransf.List(0)
  MarcarTexto CPorRetTransf
End Sub

Private Sub CPorRetServ_GotFocus()
  CPorRetServ.Clear
  CPorRetServ.AddItem "0 %"
  CPorRetServ.AddItem "30 %"
  CPorRetServ.AddItem "70 %"
  CPorRetServ.AddItem "100 %"
  CPorRetServ.Text = CPorRetServ.List(0)
  MarcarTexto CPorRetServ
End Sub

Private Sub CPorRetTransf_LostFocus()
CodPorRetTrans = 0
If Val(CPorRetTransf) = 70 Then MsgBox "El porcentaje 70% no corresponde al tipo de transaccion"
  If CPorRetTransf.ListIndex < 0 Then
     CodPorRetTrans = 0
  Else
     CodPorRetTrans = CPorRetTransf.ListIndex
  End If
  PorRetTran = Val(Mid(CPorRetTransf.Text, 1, 5))
End Sub

Private Sub DCCompMod_GotFocus()
  SelectDBCombo DCCompMod, AdoConcepto, SQL2, "Detalle"
End Sub

Private Sub DCCompMod_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComprobante_GotFocus()
  MarcarTexto DCComprobante
End Sub

Private Sub DCComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCComprobante_LostFocus()
td = SinEspaciosIzq(DCComprobante.Text)
''MsgBox td & TipoDoc
ErrorST (td & TipoDoc)
 Select Case td
 Case "04", "05"
   Label35.Visible = True
   DCCompMod.Visible = True
   Label34.Visible = True
   DCNumMod.Visible = True
   DCCompMod.SetFocus
 Case Else
   Label35.Visible = False
   DCCompMod.Visible = False
   Label34.Visible = False
   DCNumMod.Visible = False
 End Select
SQL1 = "SELECT Comprobante " _
     & "FROM Trans_Retenciones  " _
     & "WHERE Item = '" & NumEmpresa & "' " _
     & "AND TD = '" & td & "' " _
     & "AND Codigo = '" & CodigoCliente & "' " _
     & "AND Secuencial = '" & TextSec & "' " _
     & "AND Autorizacion = '" & TextAuto & "' "
 SelectAdodc AdoRetencion, SQL1
 If AdoRetencion.Recordset.RecordCount > 0 Then MsgBox "Esta transaccion ya existe"
End Sub

Private Sub DCConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCNumMod_GotFocus()
  sSQL = "SELECT Secuencial " _
       & "FROM Trans_Retenciones "
  Select Case Topc
         Case "1":  sSQL = sSQL & " WHERE TT = 'C' "
         Case "2":  sSQL = sSQL & " WHERE TT = 'V' "
         Case "3":  sSQL = sSQL & " WHERE TT = 'TI' "
         Case "4":  sSQL = sSQL & " WHERE TT = 'TE' "
  End Select
  sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
              & "AND TD <> '04' " _
              & "AND TD <> '05' " _
              & "GROUP BY Secuencial "
  SelectDBCombo DCNumMod, AdoConcepto, sSQL, "Secuencial"
End Sub

Private Sub DCNumMod_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSCT_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC IN ('RF','IV') " _
       & "AND DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY TC,Codigo "
  SelectAdodc AdoRetenido, sSQL
  CCtaRet.Clear
  CCtaIVA30.Clear
  CCtaIVA70.Clear
  With AdoRetenido.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Select Case .Fields("TC")
            Case "IV": CCtaIVA30.AddItem .Fields("Codigo") & " - " & .Fields("Cuenta")
                       CCtaIVA70.AddItem .Fields("Codigo") & " - " & .Fields("Cuenta")
            Case "RF": CCtaRet.AddItem .Fields("Codigo") & " - " & .Fields("Cuenta")
          End Select
         .MoveNext
       Loop
   End If
  End With
  CCtaIVA30.Text = CCtaIVA30.List(0)
  CCtaIVA70.Text = CCtaIVA70.List(0)
  CCtaRet.Text = CCtaRet.List(0)
  ComboCT.Clear    ' Combo credito tributario
  ComboCT.AddItem "00   No Aplica"
  ComboCT.AddItem "01   Crédito tributario para la declaración del IVA"
  ComboCT.AddItem "02   Costo o Gasto para la declaración de IR"
  ComboCT.AddItem "03   Activo Fijo - Crédito tributario para declaración de IVA"
  ComboCT.AddItem "04   Activo Fijo - Costo o Gasto para la declaración de IR"
  ComboCT.AddItem "05   Liquidación Gastos Viaje, hospedaje y alimentación Gastos IR"
  ComboCT.AddItem "06   Inventario - Crédito Tributario para la declaración de IVA"
  ComboCT.AddItem "07   Inventario - Costo o Gasto para declaración de IR"
  CPorcIVA.Clear   ' Porcentajes de iva
  CPorcIVA.AddItem "0  %"
  CPorcIVA.AddItem "10 %"
  CPorcIVA.AddItem "12 %"
  CPorcIVA.AddItem "14 %"
  CPorcRet.Clear   ' Porcentajes de retencion en la fuente
  CPorcRet.AddItem "1 %"
  CPorcRet.AddItem "5 %"
  CPorcRet.AddItem "8 %"
  ComboICE.Clear
  ComboICE.AddItem "00.00 %"
  ComboICE.AddItem "18.54 % Cigarrillos Negros "
  ComboICE.AddItem "30.90 % Cerveza"
  ComboICE.AddItem "10.30 % Bebidas Gaseosas"
  ComboICE.AddItem "15.00 % Servicios de Telecomunicaciones y Radioelectricos"
  ComboICE.AddItem " 5.15 % Vehiculos motorizados de transporte terrestre de hasta 3.5 toneladas de carga"
  ComboICE.AddItem "10.30 % Aviones, avionetas y helicópteros, motos acuaticas, tricares, cuadrones, yates y barcos de recreo"
  'MsgBox Anio
  sSQL = "SELECT (Codigo  & '    ' &  Tipo_Ret) As Detalle " _
       & "FROM Tipo_Reten " _
       & "WHERE Item = '000' AND Año = '" & Anio & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConcepto, AdoConcepto, sSQL, "Detalle"
  If IngMensual Then
      Cadena1 = Mid(Periodo, 4, 2)
      Cadena = MesesLetras(Val(Cadena1))
      Cadena = Cadena & " de " & Mid(Periodo, 7, 4)
      LblPeriodo.Caption = "Periodo:" & Cadena
  Else
      Cadena = Mid(Periodo, 7, 4)
  End If
  CmdGrabar.Enabled = False
  IvaTotal = 0
  BaseServ = 0
  BaseTrans = 0
  I = 0: J = 0
  sSQL = "SELECT Codigo,CodigoTR,Secuencial,Autorizacion,Serie,FechaC,Item,Periodo " _
       & "FROM Trans_Retenciones " _
       & "WHERE T <> 'A' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodigoTR,Secuencial "
  SelectAdodc AdoRetenido, sSQL
  With AdoRetenido.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If .Fields("CodigoTR") = "RF" Then
              If I < .Fields("Secuencial") Then I = .Fields("Secuencial")
          End If
          If .Fields("CodigoTR") = "IV" And .Fields("Codigo") = CodigoCliente Then
              If J < .Fields("Secuencial") Then J = .Fields("Secuencial")
              TextAuto = .Fields("Autorizacion")
              TextSerie = .Fields("Serie")
              MBoxFechaC = .Fields("FechaC")
          End If
         .MoveNext
       Loop
   End If
  End With
  TextNumCR = I + 1
  TextSec = J + 1
  LblProv.Caption = " " & NombreCliente & Space(60 - Len(NombreCliente)) & TipoDoc & CodigoCliente
  MBoxFechaE = FechaComp
  MBoxFechaR = FechaComp
  RatonNormal
  Opctrans(0).SetFocus
End Sub

Private Sub Form_Load()
  RatonNormal
  CentrarForm FRetIVAF
  ConectarAdodc AdoRetenido
  ConectarAdodc AdoConcepto
  ConectarAdodc AdoSCT
  ConectarAdodc AdoComprobante
  ConectarAdodc AdoRetencion
End Sub

Private Sub MBoxFechaC_GotFocus()
  MarcarTexto MBoxFechaC
End Sub

Private Sub MBoxFechaC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaC_LostFocus()
  FechaValida MBoxFechaC
End Sub

Private Sub MBoxFechaE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaE_GotFocus()
  MarcarTexto MBoxFechaE
End Sub

Private Sub MBoxFechaE_LostFocus()
  FechaValida MBoxFechaE
  Anio = Year(MBoxFechaE)
  'MsgBox Anio
  sSQL = "SELECT (Codigo  & '    ' &  Tipo_Ret) As Detalle " _
       & "FROM Tipo_Reten " _
       & "WHERE Item = '000' AND Año = '" & Anio & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConcepto, AdoConcepto, sSQL, "Detalle"
End Sub

Private Sub MBoxFechaR_GotFocus()
   MarcarTexto MBoxFechaR
End Sub

Private Sub MBoxFechaR_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaR_LostFocus()
  FechaValida MBoxFechaE
  FechaValida MBoxFechaR
  I = CFechaLong(MBoxFechaE)
  J = CFechaLong(MBoxFechaR)
  If I > J Then
      MsgBox "La fecha ingresada no es válida, debe ser Igual o Superior a la fecha de emisión del comprobante"
      MBoxFechaR = MBoxFechaE
      MBoxFechaE.SetFocus
  Else
   ComboCT.Clear
   ComboCT.AddItem "00   No Aplica"
   Cadena = BuscarFecha("01/01/2006") 'MsgBox Periodo & Cadena
   If BuscarFecha(MBoxFechaE.Text) >= Cadena Then
      ComboCT.AddItem "01   Compras netas de servicios y bienes distintos de inventarios y activos fijos que sustentan crédito tributario"
      ComboCT.AddItem "02   Compras netas de servicios y bienes distintos de inventarios y activos fijos que NO sustentan crédito tributario"
      ComboCT.AddItem "03   Compras netas de activos fijos que sustentan crédito tributario"
      ComboCT.AddItem "04   Compras netas de activos fijos que NO sustentan crédito tributario"
      ComboCT.AddItem "05   Liquidación de gastos de viaje a nombre de empleados y no de la empresa"
      ComboCT.AddItem "06   Compras netas de inventarios que sustentan crédito tributario"
      ComboCT.AddItem "07   Compras netas de inventarios que NO sustentan crédito tributario"
      ComboCT.AddItem "08   Valor pagado o recibido por Reembolso de gasto"
      ComboCT.AddItem "09   Reembolso por gastos médicos y medicina prepagada"
   Else
      ComboCT.AddItem "01   Crédito tributario para la declaración del IVA"
      ComboCT.AddItem "02   Costo o Gasto para la declaración de IR"
      ComboCT.AddItem "03   Activo Fijo - Crédito tributario para declaración de IVA"
      ComboCT.AddItem "04   Activo Fijo - Costo o Gasto para la declaración de IR"
      ComboCT.AddItem "05   Liquidación Gastos Viaje, hospedaje y alimentación Gastos IR"
      ComboCT.AddItem "06   Inventario - Crédito Tributario para la declaración de IVA"
      ComboCT.AddItem "07   Inventario - Costo o Gasto para declaración de IR"
   End If
   ComboCT.Text = ComboCT.List(0)
  End If
  'MBoxFechaCR = MBoxFechaR
End Sub

Private Sub OpcD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Opctrans_Click(Index As Integer)
  Topc = (Index + 1)
End Sub

Private Sub Opctrans_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub Opctrans_LostFocus(Index As Integer)
 CodPorRetSer = "0"
 CodPorRetTrans = "0"
 CodPorcIva = "0"
 Topc = (Index + 1)
 sSQL = "SELECT (Codigo & '  ' & Detalle) As Viste " _
      & "FROM Tipo_IVA "
 SQL2 = "SELECT (Codigo  & '    ' &  Detalle) As Detalle " _
      & "FROM Tipo_IVA "
 Label17.Width = 1290
 Label17.Caption = "Monto IVA en Prest. Servicios"
 Text3.Visible = False
 TBaseServ.Visible = True
 Label35.Visible = False
 DCCompMod.Visible = False
 Label34.Visible = False
 DCNumMod.Visible = False
 FrmPres.Visible = True
 FrmPres.Caption = "Presuntivo"
 OptP(0).Caption = "Si"
 OptP(1).Caption = "No"
 FrmPres.Visible = False
FrmDev.Visible = True  '
Label36.Visible = False  ' numero de refrendo
TextAdu.Visible = False
Label15.Visible = True ' porc retencion transferencia
CPorRetTransf.Visible = True
Label14.Visible = True  ' valor retenido en transferencia
TValRetTransf.Visible = True
Label1.Visible = True ' fecha emision
MBoxFechaE.Visible = True
Label12.Caption = "Fecha Registro Comprobante"
Label8.Visible = True 'numero serie
TextSerie.Visible = True
Label21.Visible = True ' numero secuencial
Label21.Caption = "Factura No"
TextSec.Visible = True
Label18.Visible = True 'Numero autorizacion
TextAuto.Visible = True
Label16.Visible = True ' porc ret servicios
CPorRetServ.Visible = True
Label19.Visible = True ' valor retenido servicios
TValRetServ.Visible = True
Label6.Visible = True 'valor iva
TxtValorIVA.Visible = True
Label9.Visible = True ' porc iva
CPorcIVA.Visible = True
Label10.Visible = True ' base imponible cero
Txtbasecero.Visible = True
Label13.Caption = "Base Imponible ICE" ' monto de ice
Label26.Visible = True 'porcentaje de ice
ComboICE.Visible = True
Label27.Visible = True ' VALOR RET ICE
TextValorICE.Visible = True
Label17.Caption = "Base Prestacion de Servicios"
DCNumMod.Text = "0"
ComboCT.Enabled = True
DCSCT.Enabled = True
    Select Case Topc
        Case 1 ' compras
            sSQL = sSQL & "WHERE Cod = 'TC' "
            SQL2 = SQL2 & "WHERE Cod = 'C' "
            FRetIVAF.Caption = "Compras y Retencion en la Fuente"
'            Label32.Visible = False
            FrmPres.Visible = False
            LProCli.Caption = "Proveedor:"
            Label25.Visible = True 'fecha caducidad
            MBoxFechaC.Visible = True
            Label6.Caption = "Base IVA Transf Bienes"
            ComboCT.Text = ComboCT.List(1)
        Case 2  ' ventas
             sSQL = sSQL & "WHERE Cod = 'TV' "
             SQL2 = SQL2 & "WHERE Cod = 'V' "
             FRetIVAF.Caption = "Ventas y Retención Presuntiva"
             FrmPres.Visible = True
             Label8.Visible = False
             TextSerie.Visible = False
             Label18.Visible = False 'Numero autorizacion
             TextAuto.Visible = False
             LProCli.Caption = "Cliente:"
             FrmPres.Visible = True
             Label21.Caption = "Num Facturas" ' secuencial
             FrmDev.Visible = False
             Label25.Visible = False 'fecha caducidad
             MBoxFechaC.Visible = False
             MBoxFechaC.Text = "01/01/1995"
             Label8.Visible = False 'numero serie
             TextSerie.Visible = False
             Label18.Visible = False 'Numero autorizacion
             TextAuto.Visible = False
             Label6.Caption = "Base IVA Transf Bienes"
             ComboCT.Text = "00     No Aplica"
             ComboCT.Enabled = False
        Case 3  ' importaciones
             sSQL = sSQL & "WHERE Cod = 'IT' "
             SQL2 = SQL2 & "WHERE Cod = 'TI' "
             FrmDev.Visible = False
             FrmPres.Visible = True
             FrmPres.Caption = "Importacion"
             OptP(0).Caption = "Bienes"
             OptP(1).Caption = "Servicios"
             LProCli.Caption = "Proveedor:"
             Text3.Visible = False
             Label1.Visible = False ' fecha emision
             MBoxFechaE.Visible = False
             MBoxFechaE = "01/01/1995"
             Label12.Caption = "Fecha Pago Liquidación"
             Label25.Visible = False 'fecha caducidad
             MBoxFechaC.Text = "01/01/1999"
             MBoxFechaC.Visible = False
             Label8.Visible = False 'numero serie
             TextSerie.Visible = False
             Label21.Visible = False ' numero secuencial
             TextSec.Visible = False
             Label18.Visible = False 'Numero autorizacion
             TextAuto.Visible = False
             Label36.Visible = True ' numero de refrendo
             TextAdu.Visible = True
             Label15.Visible = False 'porc retencion transferencia de bienes
             CPorRetTransf.Visible = False
             Label14.Visible = False  ' retenido en transf bienes
             TValRetTransf.Visible = False
             Label16.Visible = False ' porc ret servicios
             CPorRetServ.Visible = False
             Label19.Visible = False ' valor retenido servicios
             TValRetServ.Visible = False
             Label6.Caption = "Valor IVA Importación"
             Label17.Caption = "Valor CIF Importacion"
        Case 4
              sSQL = sSQL & "WHERE Cod = 'ET' "
              SQL2 = SQL2 & "WHERE Cod = 'TE' "
             LProCli.Caption = "Cliente:"
             FrmPres.Visible = True
             FrmPres.Caption = "Exportacion"
             OptP(0).Caption = "Bienes"
             OptP(1).Caption = "Servicios"
             TBaseServ.Visible = False
             Label17.Width = 1500
             Label17.Caption = "# Transporte"
             Text3.Visible = True
             Label36.Visible = True ' numero de refrendo
             TextAdu.Visible = True
             Label15.Visible = False ' porc ret transferencia
             CPorRetTransf.Visible = False
             Label14.Visible = False ' valor retenido tranferencia
             TValRetTransf.Visible = False
             Label25.Caption = "Fecha de Embarque"
             Label25.Visible = True 'fecha caducidad
             MBoxFechaC.Visible = True
             Label16.Visible = False ' porc ret servicios
             CPorRetServ.Visible = False
             Label19.Visible = False ' valor retenido servicios
             TValRetServ.Visible = False
             Label6.Visible = False 'valor iva
             TxtValorIVA.Visible = False
             Label9.Visible = False ' porc iva
             CPorcIVA.Visible = False
             Label10.Visible = False ' base imponible cero
             Txtbasecero.Visible = False
             Label26.Visible = False 'porcentaje de ice
             ComboICE.Visible = False
             Label27.Visible = False ' VALOR RET ICE
             TextValorICE.Visible = False
             Label13.Caption = "Valor FOB" ' monto de ice
             ComboICE.Text = ComboICE.List(0)
    End Select
sSQL = sSQL & "AND Item = '000' AND Año = '" & Anio & "' " _
     & "ORDER BY Pos "
SelectDBCombo DCSCT, AdoSCT, sSQL, "Viste", False
SQL2 = SQL2 & "AND Item = '000' AND Año = '" & Anio & "' " _
     & "ORDER BY Pos "
SelectDBCombo DCComprobante, AdoComprobante, SQL2, "Detalle"
End Sub

Private Sub OptP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TBaseServ_GotFocus()
   BaseServ = IvaTotal - BaseTrans
   Select Case Topc
   Case "1", "2"
     If BaseServ > 0 Then
       TBaseServ.Text = BaseServ
     Else
       TBaseServ = 0
     End If
   End Select
  MarcarTexto TBaseServ
End Sub

Private Sub TBaseServ_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TBaseServ_LostFocus()
TBaseServ = Format(TBaseServ.Text, "#,##0.00")
BaseServt = TBaseServ.Text
Parcial = BaseServt + BaseTrans
Select Case Topc
Case 1, 2
     If BaseServt <> BaseServ And Parcial <> IvaTotal Then
        MsgBox "El valor introducido no es correcto"
        Parcial = Parcial - BaseServt
        TBaseServ.SetFocus
     Else
        BaseServ = TBaseServ.Text
     End If
Case 3
     If (Val(TxBaseImpo) + Val(Txtbasecero)) > 0 Then CmdSave.Enabled = True
End Select
End Sub

Private Sub Text3_GotFocus()
   MarcarTexto Text3
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Text3_LostFocus()
   If Val(Text3.Text) > 0 Then
   Text3 = Format(Text3, "0000000000000")
Else
   MsgBox "El Numero de Transporte debe ser mayor que 1"
   Text3 = "0000000000001"
   Text3.SetFocus
End If
End Sub

Private Sub TextAdu_GotFocus()
   MarcarTexto TextAdu
End Sub

Private Sub TextAdu_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextAdu_LostFocus()
   If Val(TextAdu.Text) > 0 Then
   TextAdu = Format(TextAdu, "0000000000000000")
Else
   MsgBox "El Numero de Refrendo debe ser mayor que 1"
   TextAuto = "0000000000000001"
   TextAuto.SetFocus
End If
End Sub

Private Sub TextAuto_GotFocus()
   MarcarTexto TextAuto
End Sub

Private Sub TextAuto_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextAuto_LostFocus()
If Val(TextAuto.Text) > 0 Then
   TextAuto = Format(TextAuto, "0000000000")
Else
   MsgBox "El Numero de Autorizacion debe ser mayor que 1"
   TextAuto = "0000000001"
   TextAuto.SetFocus
End If
End Sub

Private Sub TextIVA_GotFocus()
    TextIVA = IvaTotal
    MarcarTexto TextIVA
End Sub

Private Sub TextIVA_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextIVA_LostFocus()
   TextIVA = Format(TextIVA, "#,##0.00")
End Sub

Private Sub TextNumCR_GotFocus()
   MarcarTexto TextNumCR
End Sub

Private Sub TextNumCR_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNumCR_LostFocus()
If Val(TextNumCR.Text) > 0 Then
   TextNumCR = Format(TextNumCR, "0000000")
Else
   MsgBox "El Numero de comprobante de Retencion debe ser mayor que 1"
   TextNumCR = "0000001"
   TextNumCR.SetFocus
End If
End Sub

Private Sub TextSec_GotFocus()
   MarcarTexto TextSec
End Sub

Private Sub TextSec_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextSec_LostFocus()
If Val(TextSec.Text) > 0 Then
   TextSec = Format(TextSec, "0000000")
Else
   MsgBox "El numero de Comprobante siempre debe ser mayor que 1"
   TextSec = "0000001"
   TextSec.SetFocus
End If
End Sub

Private Sub TextSerie_GotFocus()
   MarcarTexto TextSerie
End Sub

Private Sub TextSerie_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextSerie_LostFocus()
If Val(TextSerie.Text) > 1000 Then
   TextSerie = Format(TextSerie, "000000")
Else
   MsgBox "El Numero de Serie debe ser mayor que 001001"
   TextSerie = "001001"
   TextSerie.SetFocus
End If
End Sub

Private Sub TextValorICE_GotFocus()
   TextValorICE = TxtICE * Val(SinEspaciosIzq(ComboICE.Text)) / 100
   MarcarTexto TextValorICE
End Sub

Private Sub TextValorICE_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextValorICE_LostFocus()
   TextValorICE = Format(TextValorICE, "#,##0.00")
End Sub

Private Sub TValRetServ_GotFocus()
  Select Case Topc
 Case "1", "2"
    TValRetServ.Text = IvaServ
Case "3", "4"
     TValRetServ.Text = "0.00"
End Select
MarcarTexto TValRetServ
End Sub

Private Sub TValRetServ_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TValRetServ_LostFocus()
    TValRetServ.Text = Format(Val(TValRetServ.Text), "#,##0.00")
    IvaServ = TValRetServ.Text
    If (Val(TxBaseImpo) + Val(Txtbasecero)) > 0 Then CmdGrabar.Enabled = True
End Sub

Private Sub TValRetTransf_GotFocus()
  IvaTrans = BaseTrans * PorRetTran / 100
  TValRetTransf.Text = IvaTrans
  TValRetTransf.Text = Format(Val(TValRetTransf.Text), "#,##0.00")
  IvaTrans = Format(Val(TValRetTransf.Text), "#,##0.00")
  MarcarTexto TValRetTransf
End Sub

Private Sub TValRetTransf_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TValRetTransf_LostFocus()
   IvaTrans = TValRetTransf.Text
   TValRetTransf = Format(TValRetTransf, "#,##0.00")
End Sub

Private Sub TxBaseImpo_GotFocus()
   MarcarTexto TxBaseImpo
End Sub

Private Sub TxBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxBaseImpo_LostFocus()
   TextoValido TxBaseImpo, True
   TxBaseImpo = Format(TxBaseImpo, "#,##0.00")
   BaseImpo = TxBaseImpo.Text
End Sub

Private Sub TxBaseImpoF_GotFocus()
   MarcarTexto TxBaseImpoF
End Sub

Private Sub TxBaseImpoF_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxBaseImpoF_LostFocus()
   TextoValido TxBaseImpoF, True
   TxBaseImpoF = Format(TxBaseImpoF, "#,##0.00")
   BaseImpoF = TxBaseImpoF.Text
End Sub

Private Sub Txtbasecero_GotFocus()
  MarcarTexto Txtbasecero
End Sub

Private Sub Txtbasecero_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Txtbasecero_LostFocus()
   BaseCero = Val(Txtbasecero)
   TxtVRet = Format(Val(CCur(TxBaseImpoF)) * Val(SinEspaciosIzq(CPorcRet)) / 100, "#,##0.00")
   Txtbasecero = Format(Txtbasecero, "#,##0.00")
End Sub

Private Sub TxtICE_GotFocus()
  MarcarTexto TxtICE
End Sub

Private Sub TxtICE_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtICE_LostFocus()
   TxtICE = Format(TxtICE, "#,##0.00")
End Sub

Private Sub TxtValorIVA_GotFocus()
   TxtValorIVA = Format(TxBaseImpo * Val(CPorcIVA.Text) / 100, "#,##0.00")
   MarcarTexto TxtValorIVA
End Sub

Private Sub TxtValorIVA_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtValorIVA_LostFocus()
   BaseTrans = TxtValorIVA
   If BaseTrans <= IvaTotal Then
      TxtValorIVA = Format(TxtValorIVA, "#,##0.00")
      BaseServ = IvaTotal - BaseTrans
   Else
      MsgBox "El valor introducido no es correto"
   End If
End Sub

Private Sub TxtVRet_GotFocus()
   MarcarTexto TxtVRet
End Sub

Private Sub TxtVRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtVRet_LostFocus()
   TxtVRet = Format(Val(TxtVRet), "#,##0.00")
   'MsgBox TxtVRet
End Sub
