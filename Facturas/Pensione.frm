VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form FacturasPension 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FACTURACION DE PENSIONES"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15960
   WindowState     =   1  'Minimized
   Begin VB.ListBox LstMeses 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   35
      Top             =   4620
      Width           =   16500
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
      Left            =   105
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1470
      Width           =   2115
   End
   Begin VB.Frame FrmDebito 
      BackColor       =   &H00FFC0C0&
      Height          =   960
      Left            =   105
      TabIndex        =   24
      Top             =   3255
      Visible         =   0   'False
      Width           =   11985
      Begin VB.ComboBox CTipoCta 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   10080
         TabIndex        =   28
         Text            =   "TARJETA"
         Top             =   210
         Width           =   1800
      End
      Begin VB.CheckBox CheqPorDeposito 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Depositar al Banco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9555
         TabIndex        =   33
         Top             =   525
         Width           =   2325
      End
      Begin VB.TextBox TxtCtaNo 
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   2310
         MaxLength       =   20
         TabIndex        =   30
         Text            =   "."
         Top             =   525
         Width           =   4845
      End
      Begin MSDataListLib.DataCombo DCDebito 
         Bindings        =   "Pensione.frx":0000
         DataSource      =   "AdoDebito"
         Height          =   315
         Left            =   2310
         TabIndex        =   26
         Top             =   210
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   12582912
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
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   8400
         TabIndex        =   32
         ToolTipText     =   "Formato de Fecha: DD/MM/AA"
         Top             =   525
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   12582912
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "MM/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "0"
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
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
         Left            =   9450
         TabIndex        =   27
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DEBITO AUTOMATICO"
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
         TabIndex        =   25
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DE CUENTA"
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
         TabIndex        =   29
         Top             =   525
         Width           =   2220
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CADUCIDAD"
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
         Left            =   7140
         TabIndex        =   31
         Top             =   525
         Width           =   1275
      End
   End
   Begin VB.CheckBox CheqDebito 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ingrese sus datos para el Debito Automatico"
      BeginProperty Font 
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
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2940
      Width           =   10200
   End
   Begin VB.TextBox TxtCI_RUC 
      BackColor       =   &H00000080&
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
      Left            =   10290
      MaxLength       =   13
      TabIndex        =   9
      Top             =   1155
      Width           =   1800
   End
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   4620
      TabIndex        =   3
      Top             =   735
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
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   1575
      TabIndex        =   1
      Top             =   735
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   945
      Top             =   9975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":0018
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":08F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":11CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":1872
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":1B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":1EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":22F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":3024
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":3C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":44F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pensione.frx":4DCC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "%"
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
      Left            =   14070
      TabIndex        =   43
      Top             =   2415
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   10080
   End
   Begin VB.TextBox TxtEmail 
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
      Left            =   2205
      MaxLength       =   60
      TabIndex        =   20
      Top             =   2520
      Width           =   5160
   End
   Begin VB.TextBox TxtTelefono 
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
      Left            =   8505
      MaxLength       =   10
      TabIndex        =   22
      Top             =   2520
      Width           =   1800
   End
   Begin VB.TextBox TextCI 
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
      Left            =   8505
      MaxLength       =   13
      TabIndex        =   16
      Top             =   1890
      Width           =   1800
   End
   Begin VB.TextBox TxtDireccion 
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
      Left            =   2205
      MaxLength       =   50
      TabIndex        =   18
      Top             =   2205
      Width           =   8100
   End
   Begin VB.TextBox TextRepresentante 
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
      Left            =   2205
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1890
      Width           =   4845
   End
   Begin VB.CheckBox CheqMes 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10920
      TabIndex        =   5
      Top             =   735
      Value           =   1  'Checked
      Width           =   1170
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
      Left            =   2205
      MaxLength       =   50
      TabIndex        =   11
      Top             =   1470
      Width           =   8100
   End
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "Pensione.frx":6216
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   5985
      TabIndex        =   4
      Top             =   735
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Clientes"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoListFact 
      Height          =   330
      Left            =   2415
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
      Caption         =   "ListFact"
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
      Left            =   2415
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   4620
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
      Caption         =   "Articulo"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   2415
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
      Caption         =   "Linea"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Pensione.frx":622D
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   2205
      TabIndex        =   7
      Top             =   1155
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
      Height          =   330
      Left            =   210
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
      Caption         =   "AsientoF"
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
      Left            =   4620
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   4620
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
   Begin MSAdodcLib.Adodc AdoNC 
      Height          =   330
      Left            =   210
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
      Caption         =   "NC"
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
   Begin MSAdodcLib.Adodc AdoHistoria 
      Height          =   330
      Left            =   2415
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
      Caption         =   "Historia"
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
   Begin MSAdodcLib.Adodc AdoAnticipo 
      Height          =   330
      Left            =   4620
      Top             =   5775
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Anticipo"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Matriculas"
            Object.ToolTipText     =   "Actualizar Matriculados"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Buses"
            Object.ToolTipText     =   "Actualizacion de la Facturacion de Buses"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Historia"
            Object.ToolTipText     =   "Presenta la historia del Cliente"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CarteraPDF"
                  Text            =   "En PDF"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CarteraExcel"
                  Text            =   "En Excel"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CarteraMail"
                  Text            =   "Por Mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Deuda_Pendiente"
            Object.ToolTipText     =   "Presenta la Deuda Pendiente"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recalcular_Saldos"
            Object.ToolTipText     =   "Recalcular Saldos de Facturas"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PreFacturas"
            Object.ToolTipText     =   "Insertar Prefacturacion Mensual"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NuevoCliente"
            Object.ToolTipText     =   "Insertar nuevo Beneficiario/Cliente"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpdateCliente"
            Object.ToolTipText     =   "Actualiza datos del Cliente"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LeerJS"
            Object.ToolTipText     =   "Lee JS"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   750
         Left            =   4935
         TabIndex        =   77
         Top             =   -105
         Width           =   11670
         Begin VB.TextBox TextFacturaNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9975
            TabIndex        =   78
            Text            =   "0000000000"
            Top             =   210
            Width           =   1590
         End
         Begin MSMask.MaskEdBox MBHistorico 
            Height          =   330
            Left            =   1575
            TabIndex        =   83
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
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " &Inicio Resumen"
            BeginProperty Font 
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
            TabIndex        =   82
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2940
            TabIndex        =   79
            Top             =   210
            Width           =   6945
         End
      End
   End
   Begin VB.Frame FrmFormaPago 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3060
      Left            =   105
      TabIndex        =   49
      Top             =   6825
      Width           =   16500
      Begin VB.TextBox TextCheqNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11130
         MaxLength       =   10
         TabIndex        =   57
         Top             =   420
         Width           =   2430
      End
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "Pensione.frx":6246
         DataSource      =   "AdoBanco"
         Height          =   420
         Left            =   2310
         TabIndex        =   55
         Top             =   420
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   741
         _Version        =   393216
         Text            =   "Banco"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TextBanco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2310
         MaxLength       =   25
         TabIndex        =   51
         Text            =   "."
         Top             =   0
         Width           =   8835
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
         Height          =   855
         Left            =   10290
         Picture         =   "Pensione.frx":625D
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1890
         Width           =   1275
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
         Height          =   855
         Left            =   8925
         Picture         =   "Pensione.frx":6B27
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1890
         Width           =   1275
      End
      Begin VB.TextBox TextInteres 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14280
         MaxLength       =   10
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   2100
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox TxtCodigoC 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   80
         Top             =   1785
         Width           =   1590
      End
      Begin VB.TextBox TxtNC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   420
         Left            =   14280
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   67
         Text            =   "Pensione.frx":6F69
         Top             =   1260
         Width           =   2220
      End
      Begin VB.TextBox TxtEfectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   420
         Left            =   14280
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   69
         Text            =   "Pensione.frx":6F70
         Top             =   1680
         Width           =   2220
      End
      Begin VB.TextBox TxtSaldoFavor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   420
         Left            =   14280
         MaxLength       =   14
         TabIndex        =   63
         Text            =   "0.00"
         Top             =   840
         Width           =   2220
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   14280
         MaxLength       =   14
         TabIndex        =   59
         Text            =   "0.00"
         Top             =   420
         Width           =   2220
      End
      Begin MSDataListLib.DataCombo DCAnticipo 
         Bindings        =   "Pensione.frx":6F77
         DataSource      =   "AdoAnticipo"
         Height          =   420
         Left            =   2310
         TabIndex        =   61
         Top             =   840
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   741
         _Version        =   393216
         Text            =   "Anticipo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DCNC 
         Bindings        =   "Pensione.frx":6F91
         DataSource      =   "AdoNC"
         Height          =   420
         Left            =   2310
         TabIndex        =   65
         Top             =   1260
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   741
         _Version        =   393216
         Text            =   "Nota de Credito"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Doc. No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9870
         TabIndex        =   56
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interés Tarjeta USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11655
         TabIndex        =   70
         Top             =   2100
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALDO USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11655
         TabIndex        =   72
         Top             =   2520
         Width           =   2640
      End
      Begin VB.Label LblCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   14280
         TabIndex        =   73
         Top             =   2520
         Width           =   2220
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   81
         Top             =   1785
         Width           =   2325
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Notas Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   64
         Top             =   1260
         Width           =   2325
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   13545
         TabIndex        =   66
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EFECTIVO USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   11655
         TabIndex        =   68
         Top             =   1680
         Width           =   2640
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bancos/Tarjetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   54
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13545
         TabIndex        =   58
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Anticipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   60
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13545
         TabIndex        =   62
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Detalle del Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   2325
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo Pendiente USD "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11130
         TabIndex        =   52
         Top             =   0
         Width           =   3165
      End
      Begin VB.Label LblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99,999,999,99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   14280
         TabIndex        =   53
         Top             =   0
         Width           =   2220
      End
   End
   Begin MSAdodcLib.Adodc AdoDebito 
      Height          =   330
      Left            =   7035
      Top             =   4830
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Debito"
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
   Begin VB.Image ImgFoto 
      Height          =   1725
      Left            =   10395
      Picture         =   "Pensione.frx":6FA5
      Stretch         =   -1  'True
      Top             =   1575
      Width           =   1665
   End
   Begin VB.Label Label11 
      BackColor       =   &H00004080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   34
      Top             =   4305
      Width           =   16500
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIC (C)"
      BeginProperty Font 
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
      TabIndex        =   8
      Top             =   1155
      Width           =   750
   End
   Begin VB.Label Label35 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Vencimiento"
      BeginProperty Font 
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
      TabIndex        =   2
      Top             =   735
      Width           =   1800
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Emisióin"
      BeginProperty Font 
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
      Width           =   1485
   End
   Begin VB.Label LabelTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   48
      Top             =   3780
      Width           =   2220
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Facturado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      TabIndex        =   47
      Top             =   3780
      Width           =   2220
   End
   Begin VB.Label LabelIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   46
      Top             =   3045
      Width           =   2220
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " I.V.A. 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   12180
      TabIndex        =   45
      Top             =   3045
      Width           =   2220
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desc. x P.P."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      TabIndex        =   42
      Top             =   2415
      Width           =   1905
   End
   Begin VB.Label LabelDescuento2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   44
      Top             =   2415
      Width           =   2220
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000C0C0&
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
      Height          =   330
      Left            =   105
      TabIndex        =   19
      Top             =   2520
      Width           =   2115
   End
   Begin VB.Label LabelDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   41
      Top             =   1890
      Width           =   2220
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   8190
      TabIndex        =   15
      Top             =   1890
      Width           =   330
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
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
      Left            =   7350
      TabIndex        =   21
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
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
      Left            =   7035
      TabIndex        =   14
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION"
      BeginProperty Font 
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
      Top             =   2205
      Width           =   2115
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Descuentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      TabIndex        =   40
      Top             =   1890
      Width           =   2220
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " RAZON SOCIAL"
      BeginProperty Font 
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
      Top             =   1890
      Width           =   2115
   End
   Begin VB.Label LabelConIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   39
      Top             =   1260
      Width           =   2220
   End
   Begin VB.Label LabelSubTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14385
      TabIndex        =   37
      Top             =   735
      Width           =   2220
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      TabIndex        =   36
      Top             =   735
      Width           =   2220
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Tarifa 12%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      TabIndex        =   38
      Top             =   1260
      Width           =   2220
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CLIENTE"
      BeginProperty Font 
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
      TabIndex        =   6
      Top             =   1155
      Width           =   2115
   End
End
Attribute VB_Name = "FacturasPension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://erp.diskcoversystem.com/~ronramiro/maristas/students.php
'http://localhost/diskcoversystem/php/vista/consultarEstudiante.php?id=17572304101791758226001010
'"{""relational_data"": [{""years"":[{""grade"":[{""show"":'segundo grado'},{""show"":'Educacion General'}]}]}]}"
Dim PorCodigo As Boolean
Dim Actualiza_Cliente As Boolean
Dim Actualiza_Buses As Boolean
Dim Si_No_local As Boolean
Dim Total_Saldo_Pendiente As Currency
Dim Rubros_Facturar() As String
Dim Cta_Ant_Cli As String

Dim tempRepresentante As String
Dim tempCI As String
Dim tempTD As String
Dim tempTelefono As String
Dim tempDirS As String
Dim tempDireccion As String
Dim tempEmail As String
Dim tempGrupo As String
Dim tempCtaNo As String
Dim tempTipoCta As String
Dim tempDocumento As String
Dim tempCaducidad As String

Dim AdoDBTemp As ADODB.Recordset
 
Public Function ReadTextFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim Handle As Integer
   If LenB(Dir$(sFilePath)) > 0 Then
   
      Handle = FreeFile
      Open sFilePath For Binary As #Handle
      ReadTextFile = Space$(LOF(Handle))
      Get #Handle, , ReadTextFile
      Close #Handle
      
   End If
   
End Function

Public Sub Actualiza_Datos_Cliente()
    Documento = 0
    If AdoDebito.Recordset.RecordCount > 0 Then
       AdoDebito.Recordset.MoveFirst
       AdoDebito.Recordset.Find ("Descripcion = '" & DCDebito & "' ")
       If Not AdoDebito.Recordset.EOF Then Documento = AdoDebito.Recordset.fields("Codigo")
    End If
    Titulo = "Formulario de Actualizacion"
    If tempRepresentante <> TextRepresentante Or _
       tempCI <> TextCI Or _
       tempTD <> Label18.Caption Or _
       tempTelefono <> TxtTelefono Or _
       tempDireccion <> TxtDireccion Or _
       tempDirS <> TxtDirS Or _
       tempEmail <> TxtEmail Or _
       tempGrupo <> CGrupo.Text Or _
       tempCtaNo <> TxtCtaNo Or _
       tempTipoCta <> CTipoCta Or _
       tempDocumento <> Documento Or _
       tempCaducidad <> MBFecha Then
       Mensajes = "DESEA ACTUALIZAR DATOS DEL REPRESENTANTE"
       If BoxMensaje = vbYes Then
          sSQL = "SELECT " & Full_Fields("Clientes_Matriculas") & " " _
               & "FROM Clientes_Matriculas " _
               & "WHERE Codigo = '" & FA.CodigoC & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' "
          Select_Adodc AdoAux, sSQL
          With AdoAux.Recordset
           If .RecordCount <= 0 Then
               SetAdoAddNew "Clientes_Matriculas"
               SetAdoFields "T", Normal
               SetAdoFields "Codigo", FA.CodigoC
               SetAdoFields "Grupo_No", CGrupo.Text
               SetAdoFields "Lugar_Trabajo_R", TxtDireccion
               SetAdoFields "Representante", TextRepresentante
               SetAdoFields "Cedula_R", TextCI
               SetAdoFields "TD", Label18.Caption
               SetAdoFields "Telefono_R", TxtTelefono
               SetAdoFields "Cta_Numero", TxtCtaNo
               SetAdoFields "Tipo_Cta", CTipoCta
               SetAdoFields "Caducidad", UltimoDiaMes("01/" & MBFecha)
               SetAdoFields "Por_Deposito", CBool(CheqPorDeposito.value)
               SetAdoFields "Cod_Banco", Documento
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
           Else
              .fields("T") = Normal
              .fields("Representante") = TextRepresentante
              .fields("Grupo_No") = CGrupo.Text
              .fields("Cedula_R") = TextCI
              .fields("TD") = Label18.Caption
              .fields("Telefono_R") = TxtTelefono
              .fields("Lugar_Trabajo_R") = TxtDireccion
              .fields("Email_R") = TxtEmail
              .fields("Cta_Numero") = TxtCtaNo
              .fields("Tipo_Cta") = CTipoCta
              .fields("Caducidad") = UltimoDiaMes("01/" & MBFecha)
              .fields("Por_Deposito") = CBool(CheqPorDeposito.value)
              .fields("Cod_Banco") = Documento
              .Update
           End If
          End With
          
          sSQL = "UPDATE Clientes " _
               & "SET Grupo = '" & CGrupo.Text & "', Direccion = '" & TxtDirS.Text & "' " _
               & "WHERE Codigo = '" & FA.CodigoC & "' "
          Ejecutar_SQL_SP sSQL
       End If
''    Else
''       MsgBox "NO SE ACTUALIZARA DATOS PORQUE USTED NO HA REALIZADO CAMBIOS DEL REPRESENTANTE", , UCase(Titulo)
    End If
End Sub

Public Sub ListaDeClientes()
    sSQL = "SELECT TOP 50 Codigo, Cliente " _
         & "FROM Clientes " _
         & "WHERE T = 'N' " _
         & "AND Codigo <> '9999999999' " _
         & "AND FA <> " & Val(adFalse) & " "
    If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
    sSQL = sSQL & "ORDER BY Cliente "
    SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
End Sub

Public Sub Grabar_FA_Pensiones()
 'Procedemos a grabar la factura
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY A_No "
  Select_Adodc AdoAsientoF, sSQL
  With AdoAsientoF.Recordset
  'MsgBox "FA Pensiones: " & .RecordCount
   If .RecordCount > 0 Then
      'Actualizamos tipos de pago
       Total_Bancos = Redondear(Val(CCur(TextCheque.Text)), 2)
       Total_Anticipo = Redondear(Val(CCur(TxtSaldoFavor)), 2)
       SubTotal_NC = Redondear(Val(CCur(TxtNC.Text)), 2)
       TotalCajaMN = Redondear(Val(CCur(TxtEfectivo.Text)), 2)
       
       Calculos_Totales_Factura FA
       FA.Tipo_PRN = "FM"
       FA.Nuevo_Doc = True
       FA.Factura = Val(TextFacturaNo)
       If Existe_Factura(FA) Then
          Titulo = "FORMULARIO DE CONFIRMACION"
          Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
                   & "Ya existe " & FA.TC & " No. " & FA.Serie & "-" & Format$(FA.Factura, "000000000") & vbCrLf & vbCrLf _
                   & "Desea Reprocesarla"
          If BoxMensaje = vbYes Then FA.Nuevo_Doc = False Else GoTo NoGrabarFA
       Else
          Factura_No = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
          If FA.Factura <> Factura_No Then
             Titulo = "FORMULARIO DE CONFIRMACION"
             Mensajes = "La " & FA.TC & " No. " & FA.Serie & "-" & Format(FA.Factura, "000000000") _
                      & ", no esta Procesada, Desea Procesarla?"
             If BoxMensaje = vbYes Then FA.Nuevo_Doc = False Else GoTo NoGrabarFA
          End If
       End If
   
       SaldoPendiente = 0
       DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
       If FA.Nuevo_Doc Then FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
       TextoFormaPago = "CONTADO"
       Total_Abonos = TotalCajaMN + Total_Bancos + SubTotal_NC + Total_Anticipo
       FA.T = Pendiente
       FA.Saldo_MN = FA.Total_MN - Total_Abonos
       FA.Porc_IVA = Porc_IVA
       FA.Cliente = NombreCliente
       TA.Recibi_de = FA.Cliente
       Cta = SinEspaciosIzq(DCBanco)
       Cta1 = SinEspaciosIzq(DCNC)
      'MsgBox Total_Abonos
      .MoveFirst
       Do While Not .EOF
          Valor = .fields("TOTAL")
          Total_Desc = .fields("Total_Desc") + .fields("Total_Desc2")
          ValorDH = Valor - Total_Desc
          Codigo = .fields("Codigo_Cliente")
          Codigo1 = .fields("CODIGO")
          Codigo2 = .fields("Mes")
          Codigo3 = .fields("HABIT")
          Anio1 = .fields("TICKET")
          ID_Reg = .fields("A_No")
          Total_Abonos = Total_Abonos - ValorDH
          If Total_Abonos >= 0 Then
             sSQL = "UPDATE Clientes_Facturacion " _
                  & "SET Valor = Valor - " & Valor & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Anio1 & "' " _
                  & "AND Codigo_Inv = '" & Codigo1 & "' " _
                  & "AND Codigo = '" & Codigo & "' " _
                  & "AND Credito_No = '" & Codigo3 & "' " _
                  & "AND Mes = '" & Codigo2 & "' "
             Ejecutar_SQL_SP sSQL
          Else
            Valor = Valor + Total_Abonos
            If Valor > 0 Then
               sSQL = "UPDATE Clientes_Facturacion " _
                    & "SET Valor = " & -Total_Abonos + Total_Desc & " " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND Periodo = '" & Anio1 & "' " _
                    & "AND Codigo_Inv = '" & Codigo1 & "' " _
                    & "AND Codigo = '" & Codigo & "' " _
                    & "AND Credito_No = '" & Codigo3 & "' " _
                    & "AND Mes = '" & Codigo2 & "' "
               Ejecutar_SQL_SP sSQL
               Total_Abonos = Total_Abonos + Total_Desc
               Valor = Valor - Total_Desc
               sSQL = "UPDATE Asiento_F " _
                    & "SET TOTAL = " & Valor & ", PRECIO = " & Valor & ", Total_Desc = 0, Total_Desc2 = 0 " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND CodigoU = '" & CodigoUsuario & "' " _
                    & "AND A_No = " & ID_Reg & " "
               Ejecutar_SQL_SP sSQL
            Else
               sSQL = "DELETE * " _
                    & "FROM Asiento_F " _
                    & "WHERE Item = '" & NumEmpresa & "' " _
                    & "AND CodigoU = '" & CodigoUsuario & "' " _
                    & "AND A_No = " & ID_Reg & " "
               Ejecutar_SQL_SP sSQL
            End If
          End If
         .MoveNext
       Loop
       
       If Not ComisionEjec Then CodigoVen = Ninguno
       
      'Grabamos el numero de factura
       Calculos_Totales_Factura FA
      ' MsgBox "............."
       Grabar_Factura FA, True
      
      'Seteos de Abonos Generales para todos los tipos de abonos
       TA.T = FA.T
       TA.TP = FA.TC
       TA.Serie = FA.Serie
       TA.Autorizacion = FA.Autorizacion
       TA.CodigoC = FA.CodigoC
       TA.Factura = FA.Factura
       TA.Fecha = FA.Fecha
       TA.Cta_CxP = FA.Cta_CxP
     
      'Abono de Factura Banco o Tarjetas
       TA.Cta = Cta
       If Len(TextBanco) <= 1 Then
          PosItem = InStr(DCBanco, " ")
          TA.Banco = UCase(TrimStrg(MidStrg(DCBanco, PosItem, Len(DCBanco))))
       Else
          TA.Banco = TextBanco & " - " & UCaseStrg(Grupo_No)
       End If
       TA.Cheque = TextCheqNo
       TA.Abono = Total_Bancos
       Grabar_Abonos TA
        
      'Abono de Factura
       TA.Cta = Cta_CajaG
       TA.Banco = "EFECTIVO MN"
       TA.Cheque = UCaseStrg(Grupo_No)
       TA.Abono = TotalCajaMN
       Grabar_Abonos TA
     
      'Forma del Abono SubTotal NC
       If SubTotal_NC > 0 Then
          SubTotal_NC = SubTotal_NC - SubTotal_IVA
          TA.Cta = Cta1
          TA.Banco = "NOTA DE CREDITO"
          TA.Cheque = "VENTAS"
          TA.Abono = SubTotal_NC
          Grabar_Abonos TA
       End If
      
      'Abonos Anticipados Cta_Ant_Cli
       TA.Cta = SinEspaciosIzq(DCAnticipo)
       If Len(TextBanco.Text) > 1 Then TA.Banco = UCaseStrg(TextBanco.Text) Else TA.Banco = "ANTICIPO PENSIONES"
       TA.Cheque = UCaseStrg(Grupo_No)
       TA.Abono = Total_Anticipo
       Grabar_Abonos TA
       
      'Forma del Abono IVA NC
       If SubTotal_IVA > 0 Then
          TA.Cta = Cta_IVA
          TA.Banco = "NOTA DE CREDITO"
          TA.Cheque = "I.V.A."
          TA.Abono = SubTotal_IVA
          Grabar_Abonos TA
       End If
     
      'Abono de Factura
       TA.T = Normal
       TA.TP = "TJ"
       TA.Cta = Cta
       TA.Cta_CxP = Cta_Tarjetas
       TA.Banco = "INTERES POR TARJETA"
       TA.Cheque = TextCheqNo
       TA.Abono = Val(TextInteres)
       TA.Recibi_de = FA.Cliente
       Grabar_Abonos TA
       
       RatonNormal
       TxtEfectivo.Text = "0.00"
      'MsgBox FA.Autorizacion
       If Len(FA.Autorizacion) >= 13 Then
          If Not No_Autorizar Then SRI_Crear_Clave_Acceso_Facturas FA, False, , True
          FA.Desde = FA.Factura
          FA.Hasta = FA.Factura
          Imprimir_Facturas_CxC FacturasPension, FA, True, False, True, True
          SRI_Generar_PDF_FA FA, True
''          RutaDestino = RutaSysBases & "\TEMP\" & FA.Autorizacion & ".pdf"
''          MsgBox RutaDestino
''          SRI_Presenta_PDF FacturasPension, RutaDestino
       Else
          Mensajes = "Facturacion Multiple"
          Titulo = "IMPRESION"
          If BoxMensaje = vbYes Then
             FA.Desde = FA.Factura
             FA.Hasta = FA.Factura
             Imprimir_Facturas_CxC FacturasPension, FA
          Else
             Imprimir_Facturas FA
          End If
         'Imprimir_Comprobante_Caja TA
       End If
       RatonReloj
       TA.Autorizacion = FA.Autorizacion
       Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
       'MsgBox TA.Factura & vbCrLf & TA.TP & vbCrLf & TA.Serie
       Facturas_Impresas FA
       
       sSQL = "SELECT * " _
            & "FROM Asiento_F " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' "
       Select_Adodc AdoAsientoF, sSQL
       TextInteres = "0.00"
       TextCheque = "0.00"
       TxtEfectivo = "0.00"
       TxtNC = "0.00"
       TxtSaldoFavor = "0.00"
       LblSaldo.Caption = "0.00"
       
      'Totales de la factura
       LabelSubTotal.Caption = "0.00"
       LabelConIVA.Caption = "0.00"
       LabelDescuento.Caption = "0.00"
       LabelDescuento2.Caption = "0.00"
       LabelIVA.Caption = "0.00"
       LabelTotal.Caption = "0.00"
       LblCambio.Caption = "0.00"
       
       ListaDeClientes
       Nuevo = False
       RatonNormal
      'MsgBox Estudiante_DBF.codest
   Else
NoGrabarFA:
       RatonNormal
       MsgBox "No se procedio a grabar el documento " & FA.TC & " No. " & FA.Serie & "-" _
            & Format(FA.Factura, "000000000") & ", revise los datos ingresados y vuelva a intentar"
   End If
  End With
End Sub

Private Sub CGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub CGrupo_LostFocus()
  If Len(TrimStrg(CGrupo.Text)) <= 1 Then CGrupo.Text = tempGrupo
  If CGrupo.Text <> tempGrupo Then
     sSQL = "SELECT Direccion " _
          & "FROM Clientes " _
          & "WHERE Grupo = '" & CGrupo.Text & "' " _
          & "AND LEN(Direccion) > 1 " _
          & "AND FA <> " & Val(adFalse) & " "
     If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
     Select_Adodc AdoAux, sSQL
     If AdoAux.Recordset.RecordCount > 0 Then TxtDirS = AdoAux.Recordset.fields("Direccion")
  End If
End Sub

Private Sub CheqDebito_Click()
    If CheqDebito.value Then
       FrmDebito.Visible = True
    Else
       FrmDebito.Visible = False
    End If
End Sub

Private Sub Command1_Click()
  Unload FacturasPension
End Sub

Private Sub Command2_Click()
Dim CtaPagoMax As String
Dim ValPagoMax As Currency
Dim SiGrabarFactura As Boolean

  SiGrabarFactura = True
  If Val(LblCambio.Caption) > 0 Then
     Titulo = "PREGUNTA DE CONFIRMACION"
     Mensajes = "Usted esta intentando grabar un documento parcialmente, de verdad desea grabarlo?"
     If BoxMensaje = vbNo Then SiGrabarFactura = False
  End If
  If SiGrabarFactura Then
     TextoValido TextRepresentante, , True
     TextoValido TxtDireccion, , True
     TextoValido TxtTelefono, , True
     TextoValido TxtEmail
    
     Titulo = "Formulario de Grabacion"
     Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
              & "La Factura No. " & TextFacturaNo
    If BoxMensaje = vbYes Then
      'Determinamos el tipos de pago de la factura
       ValPagoMax = 0
       CtaPagoMax = "1"
       If ValPagoMax <= Val(TextCheque) Then
          ValPagoMax = Val(TextCheque)
          CtaPagoMax = SinEspaciosIzq(DCBanco)
       End If
       If ValPagoMax <= Val(TxtEfectivo) Then
          ValPagoMax = Val(TxtEfectivo)
          CtaPagoMax = Cta_CajaG
       End If
       If ValPagoMax <= Val(TxtNC) Then
          ValPagoMax = Val(TxtNC)
          CtaPagoMax = SinEspaciosIzq(DCNC)
       End If
       Cta_Aux = Leer_Cta_Catalogo(CtaPagoMax)
       FA.Tipo_Pago = TipoPago
       
       Actualiza_Datos_Cliente
       
       Grabar_FA_Pensiones
       FA.Nuevo_Doc = True
       TextFacturaNo = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
       
       DCLinea.SetFocus
    End If
  End If
End Sub

'''Public Sub Actualizar_Matriculas()
'''Dim CadTemp As String
'''Dim Idx_Espe As Byte
'''Dim IdCurso As Byte
'''
'''Dim FechaIAux As String
'''Dim Idx As Long
'''Dim Valor_Desc As Currency
'''Dim Ya_Matriculo As Boolean
'''
'''   Progreso_Barra.Mensaje_Box = "ACTUALIZANDO COBROS POR GRUPO"
'''   Progreso_Iniciar
'''
'''  'Solo Matriculas
'''    With Dato_DBF
'''         FechaIAux = .FechaI
'''         sSQL = "DELETE * " _
'''              & "FROM Clientes_Facturacion " _
'''              & "WHERE Item = '" & NumEmpresa & "' " _
'''              & "AND Periodo >= '" & TrimStrg(Str(Year(.FechaI))) & "' " _
'''              & "AND Num_Mes = " & Dato_DBF.Mes_Mat & " " _
'''              & "AND Codigo_Inv IN (" & "'" & .Cod_Mat_Ini & "'," & "'" & .Cod_Mat_EBG & "'," & "'" & .Cod_Mat_Bach & "') "
'''         Ejecutar_SQL_SP sSQL
'''        'Facturamos el seguro de accidente
'''         sSQL = "DELETE * " _
'''              & "FROM Clientes_Facturacion " _
'''              & "WHERE Item = '" & NumEmpresa & "' " _
'''              & "AND Periodo >= '" & TrimStrg(Str(Year(.FechaI))) & "' " _
'''              & "AND Num_Mes = " & Dato_DBF.Mes_Mat & " " _
'''              & "AND Codigo_Inv = '01.99' "
'''         Ejecutar_SQL_SP sSQL
'''    End With
'''
'''    TextoImprimio = ""
'''
'''    sSQL = "UPDATE Clientes " _
'''         & "SET X = 'RE' " _
'''         & "WHERE FA <> " & Val(adFalse) & " "
'''    Ejecutar_SQL_SP sSQL
'''
'''    sSQL = "SELECT * " _
'''         & "FROM Clientes_Matriculas " _
'''         & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Item = '" & NumEmpresa & "' " _
'''         & "ORDER BY Codigo "
'''    Select_Adodc AdoAux, sSQL
'''
'''    sSQL = "SELECT T,Codigo,Grupo,Cliente " _
'''         & "FROM Clientes " _
'''         & "WHERE FA <> " & Val(adFalse) & " " _
'''         & "AND Grupo <> 'RETIRADO' " _
'''         & "ORDER BY Grupo,Cliente "
'''    Select_Adodc AdoCliente, sSQL
'''    With AdoCliente.Recordset
'''     If .RecordCount > 0 Then
'''         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2)
'''         Do While Not .EOF
'''           'Clientes
'''           'MsgBox ".."
'''           'If Codigo_DBF = "120150001" Then MsgBox sSQL & vbCrLf & FechaIAux & " ..."
'''            Actualiza_Cliente = False
'''            Valor_Desc = 0
'''            CodigoCli = .fields("Codigo")
'''            Grupo_No = .fields("Grupo")
'''            T = .fields("T")
'''            Progreso_Barra.Mensaje_Box = "ACTUALIZANDO COBROS DEL GRUPO: " & Grupo_No
'''            Progreso_Esperar
'''
'''           'Verificamos si ya pago por adelantado o recien los abonos de las pensiones
'''            Select Case Val(MidStrg(Grupo_No, 1, 1))
'''              Case 1
'''                   CodigoP = Dato_DBF.Cod_Mat_Ini
'''                   Valor = Dato_DBF.Val_Mat_Ini
'''              Case 2
'''                   CodigoP = Dato_DBF.Cod_Mat_EBG
'''                   Valor = Dato_DBF.Val_Mat_EBG
'''              Case 3
'''                   CodigoP = Dato_DBF.Cod_Mat_Bach
'''                   Valor = Dato_DBF.Val_Mat_Bach
'''            End Select
'''            Ya_Matriculo = True
'''            FechaIAux = Dato_DBF.FechaI
'''            sSQL = "SELECT TOP 1 Ticket,Mes_No,Total,Total_Desc,Fecha " _
'''                 & "FROM Detalle_Factura " _
'''                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''                 & "AND Item = '" & NumEmpresa & "' " _
'''                 & "AND Fecha >= #" & BuscarFecha(Dato_DBF.FechaI) & "# " _
'''                 & "AND T <> '" & Anulado & "' " _
'''                 & "AND Codigo = '" & CodigoP & "' " _
'''                 & "AND CodigoC = '" & CodigoCli & "' " _
'''                 & "ORDER BY Ticket DESC,Mes_No DESC,Fecha DESC "
'''            Select_Adodc AdoArticulo, sSQL
'''            If AdoArticulo.Recordset.RecordCount > 0 Then Ya_Matriculo = False
'''
'''           'Solo un Rubro de Matricula
'''            Progreso_Barra.Mensaje_Box = "ACTUALIZANDO COBROS DEL GRUPO: " & Grupo_No
'''            Progreso_Esperar
'''            If Ya_Matriculo Then
'''               Idx = CFechaLong(FechaIAux)
'''              'Costo Matricula
'''               SetAdoAddNew "Clientes_Facturacion"
'''               SetAdoFields "T", T
'''               SetAdoFields "Codigo", CodigoCli
'''               SetAdoFields "Codigo_Inv", CodigoP
'''               SetAdoFields "Valor", Valor
'''               SetAdoFields "Descuento", 0
'''               SetAdoFields "GrupoNo", Grupo_No
'''               SetAdoFields "Num_Mes", Dato_DBF.Mes_Mat
'''               SetAdoFields "Mes", MesesLetras(CInt(Dato_DBF.Mes_Mat))
'''               SetAdoFields "Fecha", PrimerDiaMes(FechaIAux)
'''               SetAdoFields "Periodo", CStr(Year(FechaIAux))
'''               SetAdoFields "Item", NumEmpresa
'''               SetAdoUpdate
'''            End If
'''
'''           'Facturamos el seguro de accidente
'''            Ya_Matriculo = True
'''            sSQL = "SELECT TOP 1 Ticket,Mes_No,Total,Total_Desc,Fecha " _
'''                 & "FROM Detalle_Factura " _
'''                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''                 & "AND Item = '" & NumEmpresa & "' " _
'''                 & "AND Fecha >= #" & BuscarFecha(Dato_DBF.FechaI) & "# " _
'''                 & "AND T <> '" & Anulado & "' " _
'''                 & "AND Codigo = '01.99' " _
'''                 & "AND CodigoC = '" & CodigoCli & "' " _
'''                 & "ORDER BY Ticket DESC,Mes_No DESC,Fecha DESC "
'''            Select_Adodc AdoArticulo, sSQL
'''            If AdoArticulo.Recordset.RecordCount > 0 Then Ya_Matriculo = False
'''
'''           'Solo un Rubro de Matricula
'''            Progreso_Barra.Mensaje_Box = "ACTUALIZANDO COBROS DEL GRUPO: " & Grupo_No
'''            Progreso_Esperar
'''            If Ya_Matriculo Then
'''               Idx = CFechaLong(FechaIAux)
'''              'Costo de Seguro Estudiantil
'''               SetAdoAddNew "Clientes_Facturacion"
'''               SetAdoFields "T", T
'''               SetAdoFields "Codigo", CodigoCli
'''               SetAdoFields "Codigo_Inv", "01.99"
'''               SetAdoFields "Valor", 20.58
'''               SetAdoFields "Descuento", 0
'''               SetAdoFields "GrupoNo", Grupo_No
'''               SetAdoFields "Num_Mes", Dato_DBF.Mes_Mat
'''               SetAdoFields "Mes", MesesLetras(CInt(Dato_DBF.Mes_Mat))
'''               SetAdoFields "Fecha", PrimerDiaMes(FechaIAux)
'''               SetAdoFields "Periodo", CStr(Year(FechaIAux))
'''               SetAdoFields "Item", NumEmpresa
'''               SetAdoUpdate
'''            End If
'''           .MoveNext
'''         Loop
'''     End If
'''    End With
'''    ListaDeClientes
'''    RatonNormal
'''    Progreso_Final
'''    MsgBox "Actualizacion exitosa"
'''    If Len(TextoImprimio) > 1 Then
'''       TextoImprimio = "NOMBRES DE RETIRADOS:" & vbCrLf & TextoImprimio
'''       FInfoError.Show
'''    End If
'''End Sub

Private Sub Command4_Click()
Dim Porc_Desc2 As Double
Dim SubTotal_Desc2 As Currency

Dim Valor_Desc2 As Currency
Dim ContDesc As Integer
Dim S_Porc_Desc As String

Dim S_Valor As String
Dim S_Descuento1 As String
Dim S_Descuento2 As String
Dim S_SubTotal As String

  ContDesc = 0
  SubTotal_Desc2 = CCur(LabelSubTotal.Caption) + CCur(LabelConIVA.Caption) - CCur(LabelDescuento.Caption)
  
  If SubTotal_Desc2 > 0 Then
     Valor_Desc2 = 0
     S_Porc_Desc = InputBox("Porcentaje del Descuento: ", "PORCENTAJE DE DESCUENTO", "0.00")
     If IsNumeric(S_Porc_Desc) Then
        Porc_Desc2 = Val(S_Porc_Desc)
        For I = 0 To LstMeses.ListCount - 1
            If LstMeses.Selected(I) Then ContDesc = ContDesc + 1
        Next I
        If ContDesc <> 0 Then Valor_Desc2 = Redondear(((Porc_Desc2 / 100) * SubTotal_Desc2) / ContDesc, 2)
        For I = 0 To LstMeses.ListCount - 1
         If LstMeses.Selected(I) Then
            S_Valor = TrimStrg(MidStrg(Rubros_Facturar(I), 76, 13))
            S_Descuento1 = TrimStrg(MidStrg(Rubros_Facturar(I), 89, 13))
            S_Descuento2 = Format$(Valor_Desc2, "0.00")
            S_SubTotal = Format$(Val(S_Valor) - Val(S_Descuento1) - Val(S_Descuento2), "0.00")
            Rubros_Facturar(I) = "X" & MidStrg(Rubros_Facturar(I), 1, 75) _
                               & Space(13 - Len(S_Valor)) & S_Valor _
                               & Space(13 - Len(S_Descuento1)) & S_Descuento1 _
                               & Space(13 - Len(S_Descuento2)) & S_Descuento2 _
                               & Space(13 - Len(S_SubTotal)) & S_SubTotal _
                               & "       " & SinEspaciosDer(Rubros_Facturar(I))
         End If
        Next I
        ContDesc = LstMeses.ListCount
        SubTotal_Desc2 = 0
        LstMeses.Clear
        For I = 0 To ContDesc - 1
            If MidStrg(Rubros_Facturar(I), 1, 1) = "X" Then
               SubTotal_Desc2 = SubTotal_Desc2 + Valor_Desc2
               Rubros_Facturar(I) = MidStrg(Rubros_Facturar(I), 2, Len(Rubros_Facturar(I)))
               LstMeses.AddItem Rubros_Facturar(I)
               LstMeses.Selected(I) = True
            Else
               LstMeses.AddItem Rubros_Facturar(I)
            End If
        Next I
        FA.Descuento2 = SubTotal_Desc2
        LabelDescuento2.Caption = Format$(FA.Descuento2, "#,##0.00")
     End If
  Else
     MsgBox "No tiene items a descontar"
  End If
  LstMeses.SetFocus
End Sub

Private Sub DCAnticipo_GotFocus()
  MarcarTexto DCAnticipo
End Sub

Private Sub DCAnticipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAnticipo_LostFocus()
  Total_Anticipo = Saldo_De_Anticipos(SinEspaciosIzq(DCAnticipo))
  TxtSaldoFavor = Format$(Total_Anticipo, "#,##0.00")
End Sub

Private Sub DCBanco_GotFocus()
  TextInteres = "0.00"
  TextInteres.Visible = False
  Label9.Visible = False
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBanco_LostFocus()
  TextInteres = "0.00"
  With AdoBanco.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("NomCuenta = '" & DCBanco.Text & "' ")
       If Not .EOF Then
          If .fields("TC") = "TJ" Then
              TextInteres.Visible = True
              Label9.Visible = True
          End If
       End If
   End If
  End With
End Sub

Private Sub DCLinea_GotFocus()
  Grupo_No = Ninguno
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  FA.Cod_CxC = DCLinea
  Lineas_De_CxC FA
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  TextFacturaNo = FA.Factura
  
   If FA.TC = "NV" Then
      FacturasPension.Caption = "INGRESAR NOTA DE VENTA (" & FA.TC & ")"
      Label2.Caption = " NOTA DE VENTA No. "
      Label3.Caption = " I.V.A. 0.00%"
   Else
      FacturasPension.Caption = "INGRESAR FACTURA (" & FA.TC & ")"
      Label2.Caption = " FACTURA No. "
      Label3.Caption = " I.V.A. " & Format$(Porc_IVA * 100, "#0") & "%"
   End If
   Label2.Caption = Label2.Caption & FA.Serie & "-" & FA.Autorizacion
 'MsgBox Label2.Caption & vbCrLf & TextFacturaNo
End Sub

Private Sub DCCliente_GotFocus()
  MarcarTexto DCCliente
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Codigo,Cliente " _
            & "FROM Clientes " _
            & "WHERE T = 'N' " _
            & "AND FA <> " & Val(adFalse) & " "
       If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "AND CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "AND Cliente LIKE '%" & Busqueda & "%' "
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoCliente, sSQL
    End If
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    If KeyCode = vbKeyEscape Then
       Nuevo = True
       DCBanco.SetFocus
    End If
    PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
Dim ExisteCliente As Boolean
Dim S_Valor As String
Dim S_Descuento1 As String
Dim S_Descuento2 As String
Dim S_SubTotal As String

  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  AdoAsientoF.Refresh
  DCCliente.Text = UCaseStrg(DCCliente.Text)
  CheqDebito.value = 0
  FrmDebito.Visible = False
  ExisteCliente = False
  Nuevo = False
  LstMeses.Clear
  TextRepresentante = Ninguno
  TxtEmail = Ninguno
  TextCI = Ninguno
  LblSaldo.Caption = "0.00"
  LblCambio.Caption = "0.00"
  TextInteres = "0.00"
  TxtSaldoFavor = "0.00"
  TxtEfectivo = "0.00"
  TextCheque = "0.00"
  TxtNC = "0.00"
  CodigoCliente = Ninguno
  NombreCliente = Ninguno
  DireccionCli = Ninguno
  TextRepresentante = Ninguno
  TxtDireccion = Ninguno
  TxtTelefono = Ninguno
  Total_Anticipo = 0
  Total_Saldo_Pendiente = 0
  TMail.para = ""
  DCDebito.Text = Ninguno
  CTipoCta.Text = Ninguno
  TxtCtaNo.Text = Ninguno
  CheqPorDeposito.value = 0
  MBFecha.Text = Format$(FechaSistema, "MM/yyyy")
  TBeneficiario.Saldo_Pendiente = 0
  
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCCliente & "' ")
       If Not .EOF Then
          ExisteCliente = True
          CodigoCliente = .fields("Codigo")
          FA.CodigoC = CodigoCliente
          FA = Leer_Datos_Cliente_FA(FA)
          
          tempGrupo = FA.Grupo
          CGrupo.Text = FA.Grupo
          DireccionCli = FA.DireccionC
          SQLMsg1 = DireccionCli
          SQLMsg3 = "BENEFICIARIO: " & TBeneficiario.Cliente
          
          Label10.Caption = " CLIENTE/ALUMNO (" & FA.TD & ")"
          Label18.Caption = FA.TD
          TextCI = FA.RUC_CI
          TxtCodigoC = FA.CodigoC
          TxtDirS = FA.Curso
          TxtDireccion = FA.DireccionC  '  TBeneficiario.Direccion_Rep
          TxtTelefono = FA.TelefonoC
         'MsgBox TBeneficiario.Telefono & vbCrLf & TBeneficiario.Telefono1 & vbCrLf & TBeneficiario.TelefonoT
          TextRepresentante = FA.Razon_Social ' TBeneficiario.Representante
          TxtEmail = Ninguno
          If Len(FA.EmailR) > 1 Then
             TxtEmail = FA.EmailR
          ElseIf Len(FA.EmailC) > 1 Then
             TxtEmail = FA.EmailC
          ElseIf Len(FA.EmailC2) > 1 Then
             TxtEmail = FA.EmailC2
          End If
          
          DireccionGuia = FA.DireccionC
          CodigoCliente = FA.CodigoC
          NombreCliente = FA.Cliente
          CodigoB = FA.RUC_CI
          TxtCI_RUC = FA.CI_RUC
          TelefCliente = FA.TelefonoC

          'ImgFoto.Picture = LoadPicture(RutaSistema & "\FOTOS\SINFOTO.jpg")
          
          RutaDestino = RutaSistema & "\FOTOS\" & TBeneficiario.Archivo_Foto & ".jpg"
          If Dir(RutaDestino) <> "" Then
             ImgFoto.Picture = LoadPicture(RutaDestino)
          Else
             RutaDestino = RutaSistema & "\FOTOS\" & TBeneficiario.Archivo_Foto & ".gif"
             If Dir(RutaDestino) <> "" Then ImgFoto.Picture = LoadPicture(RutaDestino)
          End If

         'Insertamos los mails de envio
          TMail.para = ""
          Insertar_Mail TMail.para, FA.EmailR
          Insertar_Mail TMail.para, FA.EmailC
          Insertar_Mail TMail.para, FA.EmailC2
          
         'Lista de Alumnos Matriculados
          CTipoCta = FA.Tipo_Cta
          Documento = FA.Cod_Banco
          TxtCtaNo = FA.Cta_Numero
          MBFecha = FA.Fecha_Cad
          If FA.Por_Deposito Then CheqPorDeposito.value = 1 Else CheqPorDeposito.value = 0
          
          If AdoDebito.Recordset.RecordCount > 0 Then
             AdoDebito.Recordset.MoveFirst
             AdoDebito.Recordset.Find ("Codigo = " & Documento & " ")
             If Not AdoDebito.Recordset.EOF Then
                DCDebito = AdoDebito.Recordset.fields("Descripcion")
             Else
                DCDebito = Ninguno
             End If
          End If
          
         'If Mod_Fact Then TextFacturaNo.SetFocus
       Else
          Nuevo = True
          MsgBox "Registro no existente"
       End If
   Else
       Nuevo = True
       MsgBox "Registro no existente"
   End If
  End With
  
  Total_Anticipo = Saldo_De_Anticipos(Cta_Ant_Cli)
  Total_Saldo_Pendiente = TBeneficiario.Saldo_Pendiente
 'MsgBox "Desktop Test: " & ExisteCliente
  If ExisteCliente Then
     If FacturaMatricula Then
        sSQL = "SELECT CF.Mes,CF.Num_Mes,CF.Valor,CF.Descuento,CF.Descuento2,CF.Codigo,CF.Periodo As Periodos,CF.Mensaje,CF.Credito_No,CP.* " _
             & "FROM Clientes_Facturacion As CF,Catalogo_Productos As CP " _
             & "WHERE CF.Codigo = '9999999999' " _
             & "AND CP.Item = '" & NumEmpresa & "' " _
             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
             & "AND CF.Mes <> '" & Ninguno & "' " _
             & "AND CF.Item = CP.Item " _
             & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
             & "ORDER BY CF.Periodo,CF.Num_Mes,CP.Codigo_Inv,CF.Credito_No "
     Else
        sSQL = "SELECT CF.Mes,CF.Num_Mes,CF.Valor,CF.Descuento,CF.Descuento2,CF.Codigo,CF.Periodo As Periodos,CF.Mensaje,CF.Credito_No,CP.* " _
             & "FROM Clientes_Facturacion As CF,Catalogo_Productos As CP " _
             & "WHERE CF.Codigo = '" & FA.CodigoC & "' " _
             & "AND CP.Item = '" & NumEmpresa & "' " _
             & "AND CP.Periodo = '" & Periodo_Contable & "' " _
             & "AND CF.Mes <> '" & Ninguno & "' " _
             & "AND CF.Item = CP.Item " _
             & "AND CF.Codigo_Inv = CP.Codigo_Inv " _
             & "ORDER BY CF.Periodo,CF.Num_Mes,CP.Codigo_Inv,CF.Credito_No "
     End If
     Select_Adodc AdoArticulo, sSQL
     With AdoArticulo.Recordset
      If .RecordCount > 0 Then
          S_Valor = "V A L O R"
          S_Descuento1 = "DESCUENTO"
          S_Descuento2 = "DESC. P.P."
          S_SubTotal = "T O T A L"
          Cadena = "  M E S" & Space(12 - Len("M E S")) _
                 & "C O D I G O" & Space(17 - Len("C O D I G O")) _
                 & "AÑO" & Space(6 - Len("AÑO")) _
                 & "P R O D U C T O" & Space(40 - Len("P R O D U C T O")) _
                 & Space(13 - Len(S_Valor)) & S_Valor _
                 & Space(13 - Len(S_Descuento1)) & S_Descuento1 _
                 & Space(13 - Len(S_Descuento2)) & S_Descuento2 _
                 & Space(13 - Len(S_SubTotal)) & S_SubTotal
          Label11.Caption = Cadena
          Do While Not .EOF
             S_Valor = Format$(.fields("Valor"), "#,##0.00")
             S_Descuento1 = Format$(.fields("Descuento"), "#,##0.00")
             S_Descuento2 = Format$(.fields("Descuento2"), "#,##0.00")
             S_SubTotal = Format$(.fields("Valor") - (.fields("Descuento") + .fields("Descuento2")), "#,##0.00")
             Producto = .fields("Producto")
             If Len(.fields("Mensaje")) > 1 Then Producto = Producto & " " & .fields("Mensaje")
             Producto = TrimStrg(MidStrg(Producto, 1, 40))
             LstMeses.AddItem .fields("Mes") & Space(12 - Len(.fields("Mes"))) _
                            & .fields("Codigo_Inv") & Space(17 - Len(.fields("Codigo_Inv"))) _
                            & .fields("Periodos") & Space(6 - Len(.fields("Periodos"))) _
                            & Producto & Space(40 - Len(Producto)) _
                            & Space(13 - Len(S_Valor)) & S_Valor _
                            & Space(13 - Len(S_Descuento1)) & S_Descuento1 _
                            & Space(13 - Len(S_Descuento2)) & S_Descuento2 _
                            & Space(13 - Len(S_SubTotal)) & S_SubTotal _
                            & "       " & .fields("Credito_No")
            .MoveNext
          Loop
          ReDim Rubros_Facturar(LstMeses.ListCount) As String
          For I = 0 To LstMeses.ListCount - 1
              Rubros_Facturar(I) = LstMeses.List(I)
          Next I
      Else
          MsgBox "No existe datos para Facturar"
          DCLinea.SetFocus
      End If
     End With
  End If
  If FA.CI_RUC = "9999999999999" Then
     FA.TD = "R"
     FA.Cliente = "CONSUMIDOR FINAL"
  End If
  Label20.Caption = " NIC(" & FA.TD & ")"
  TxtSaldoFavor = Format$(Total_Anticipo, "#,##0.00")
  LblSaldo.Caption = Format$(Total_Saldo_Pendiente, "#,##0.00")
  tempRepresentante = TextRepresentante
  tempCI = TextCI
  tempTD = Label18.Caption
  tempTelefono = TxtTelefono
  tempDireccion = TxtDireccion
  tempEmail = TxtEmail
  tempGrupo = CGrupo.Text
  tempDirS = TxtDirS
  tempCtaNo = TxtCtaNo
  tempTipoCta = CTipoCta
  tempDocumento = Documento
  tempCaducidad = MBFecha
  
  If Len(TxtCtaNo) > 1 And Len(CTipoCta) > 1 And Documento > 0 Then
     CheqDebito.value = 1
     FrmDebito.Visible = True
  End If
  
'''  If Not ExisteCliente Then
'''     TxtClienteN.Text = NombreCliente
'''     FrmPensiones.Visible = True
'''     TxtCI_RUC_N.SetFocus
'''  End If
  If Nuevo Then DCBanco.SetFocus
End Sub

Private Sub Form_Activate()
  Encerar_Factura FA
  ComisionEjec = Leer_Campo_Empresa("Comision_Ejecutivo")
  
  sSQL = "SELECT MIN(Fecha) As MinFecha " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 And Not IsNull(AdoAux.Recordset.fields("MinFecha")) Then MBHistorico = PrimerDiaMes(AdoAux.Recordset.fields("MinFecha")) Else MBHistorico = "01/01/2000"
  
  sSQL = "SELECT Codigo, Descripcion " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'BANCOS Y COOP' " _
       & "AND Codigo >= '0' " _
       & "ORDER BY Descripcion "
  SelectDB_Combo DCDebito, AdoDebito, sSQL, "Descripcion"
  
  FA.TC = TipoFactura
  Timer1.Interval = 800
 'MsgBox Leer_Campo_Empresa("Actualizar_Buses")
  Cta_Ant_Cli = Leer_Seteos_Ctas("Cta_Anticipos_Clientes")
'''  Actualiza_Buses = Leer_Campo_Empresa("Actualizar_Buses")
'''  If Actualiza_Buses Then
'''     Toolbar1.buttons("Actualizar").Enabled = False
'''     Toolbar1.buttons("Matriculas").Enabled = False
'''     Toolbar1.buttons("Pensiones").Enabled = False
'''     Toolbar1.buttons("Buses").Enabled = True
'''  Else
'''     Toolbar1.buttons("Matriculas").Enabled = True
'''     Toolbar1.buttons("Buses").Enabled = False
'''  End If
  
  If Modulo = "GERENCIA" Then Command2.Enabled = False
  
  TextFacturaNo.Enabled = Mod_Fact
  Nuevo = False
  CantFact = 1
  Contador = 0
  PorCodigo = ReadSetDataNum("PorCodigo", True, False)
  NumFacturas = ReadSetDataNum("No_FacturasImp", True, False)
  Modificar = False
  Bandera = True
  
  CTipoCta.Clear
  CTipoCta.AddItem "CORRIENTE"
  CTipoCta.AddItem "AHORROS"
  CTipoCta.AddItem "TARJETA"
  CTipoCta.AddItem Ninguno
  CTipoCta.Text = Ninguno
  
  CGrupo.Clear
  sSQL = "SELECT Grupo " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          CGrupo.AddItem .fields("Grupo")
         .MoveNext
       Loop
   End If
  End With
  
  sSQL = "SELECT (Codigo & Space(5) & Cuenta) As NomCuenta, TC, Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC IN ('BA','CJ','TJ') " _
       & "AND DG = 'D' " _
       & "ORDER BY TC,Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
   
  sSQL = "SELECT (Codigo & Space(5) & Cuenta) As NomCuenta, TC, Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC IN ('C','P') " _
       & "AND DG = 'D' " _
       & "ORDER BY TC DESC,Codigo "
  SelectDB_Combo DCAnticipo, AdoAnticipo, sSQL, "NomCuenta"
  With AdoAnticipo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo = '" & Cta_Ant_Cli & "' ")
       If Not .EOF Then
          DCAnticipo.Text = .fields("NomCuenta")
       Else
          Total_Anticipo = 0
       End If
   Else
       TxtSaldoFavor.Enabled = False
       Total_Anticipo = 0
   End If
  End With
  
  sSQL = "SELECT (Codigo & Space(5) & Cuenta) As NomCuenta, TC " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND MidStrg(Codigo,1,1) = '4' " _
       & "AND DG = 'D' " _
       & "ORDER BY TC,Codigo "
  SelectDB_Combo DCNC, AdoNC, sSQL, "NomCuenta"
  
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL

  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Select_Adodc AdoAsientoF, sSQL
   
   Label29.Caption = " Codigo Interno"
   
   ListaDeClientes
  'SeteosCtas
   If AdoCliente.Recordset.RecordCount > 0 Then Label29.Caption = " Codigo Interno (" & AdoCliente.Recordset.RecordCount & ")"
   If Bloquear_Control Then Command2.Enabled = False
   RatonNormal
   FacturasPension.WindowState = 2
   FrmFormaPago.Top = MDI_Y_Max - 3250
   LstMeses.Height = FrmFormaPago.Top - LstMeses.Top
   MBoxFecha.SetFocus
End Sub

Private Sub Form_Load()
   ConectarAdodc AdoNC
   ConectarAdodc AdoAux
   ConectarAdodc AdoCta
   ConectarAdodc AdoBanco
   ConectarAdodc AdoLinea
   ConectarAdodc AdoDebito
   ConectarAdodc AdoFactura
   ConectarAdodc AdoAsientoF
   ConectarAdodc AdoCliente
   ConectarAdodc AdoListFact
   ConectarAdodc AdoArticulo
   ConectarAdodc AdoAnticipo
   ConectarAdodc AdoHistoria
   
   SRI_Obtener_Datos_Comprobantes_Electronicos
   
   FA.CodigoC = "9999999999"
   TBeneficiario = Leer_Datos_Cliente_SP(FA.CodigoC)
   Timer1.Interval = 120
   Si_No_local = False
End Sub

Private Sub LstMeses_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Idx As Integer
Dim Porc_Desc2 As Double
Dim SubTotal_Desc2 As Currency

Dim Valor_Desc2 As Currency
Dim ContDesc As Integer
Dim S_Porc_Desc As String

Dim S_Valor As String
Dim S_Descuento1 As String
Dim S_Descuento2 As String
Dim S_SubTotal As String
  
  Idx = LstMeses.ListIndex
  Keys_Especiales Shift
  If KeyCode = vbKeyEscape Then
     Nuevo = True
     TextBanco.SetFocus
  End If
  If AltDown And KeyCode = vbKeyP Then
     If LstMeses.Selected(Idx) Then
        Valor_Desc2 = Val(TrimStrg(MidStrg(LstMeses.List(Idx), 102, 13)))
        S_Valor = TrimStrg(MidStrg(Rubros_Facturar(Idx), 76, 13))
        S_Descuento1 = TrimStrg(MidStrg(Rubros_Facturar(Idx), 89, 13))
        S_Descuento2 = Format$(Valor_Desc2, "0.00")
        Valor_Desc2 = InputBox("[" & Idx & "] Ingrese el Valor USD: ", "DESCUENTO PRONTO PAGO", S_Descuento2)
        S_Descuento2 = Format$(Valor_Desc2, "0.00")
        S_SubTotal = Format$(Val(S_Valor) - Val(S_Descuento1) - Val(S_Descuento2), "0.00")
        If Valor_Desc2 >= 0 And Valor_Desc2 <= Val(S_Valor) Then
           Rubros_Facturar(Idx) = MidStrg(Rubros_Facturar(Idx), 1, 75) _
                                & Space(13 - Len(S_Valor)) & S_Valor _
                                & Space(13 - Len(S_Descuento1)) & S_Descuento1 _
                                & Space(13 - Len(S_Descuento2)) & S_Descuento2 _
                                & Space(13 - Len(S_SubTotal)) & S_SubTotal _
                                & "       " & SinEspaciosDer(Rubros_Facturar(Idx))
           LstMeses.List(Idx) = Rubros_Facturar(Idx)
           SubTotal_Desc2 = 0
           For I = 0 To LstMeses.ListCount - 1
               SubTotal_Desc2 = SubTotal_Desc2 + Val(TrimStrg(MidStrg(Rubros_Facturar(I), 102, 13)))
           Next I
           FA.Descuento2 = SubTotal_Desc2
           LabelDescuento2.Caption = Format$(FA.Descuento2, "#,##0.00")
        Else
           MsgBox "El Valor ingresado es incorrecto"
        End If
     Else
        MsgBox "No se puede cambiar un rubro no asignado"
     End If
  End If
End Sub

Private Sub LstMeses_LostFocus()
   Factura_No = Val(TextFacturaNo.Text)
   sSQL = "DELETE * " _
        & "FROM Asiento_F " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Codigo_Cliente = '" & CodigoCliente & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   For I = 0 To LstMeses.ListCount - 1
     If LstMeses.Selected(I) Then
        MiMes = TrimStrg(SinEspaciosIzq(LstMeses.List(I)))
        Cadena = TrimStrg(MidStrg(LstMeses.List(I), Len(MiMes) + 1, Len(LstMeses.List(I))))
        Codigo = SinEspaciosIzq(Cadena)     ' Codigo_Inv
        Cadena = TrimStrg(MidStrg(Cadena, Len(Codigo) + 1, Len(Cadena)))
        Codigo1 = SinEspaciosIzq(Cadena)    ' Periodo
        Codigo2 = SinEspaciosDer(LstMeses.List(I))
        NoMeses = LetrasMeses(MiMes)
        With AdoArticulo.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
             Do While Not .EOF
                If .fields("Periodos") = Codigo1 And _
                   .fields("Num_Mes") = NoMeses And _
                   .fields("Codigo_Inv") = Codigo And _
                   .fields("Credito_No") = Codigo2 Then
                    Producto = .fields("Producto")
                    If Len(.fields("Mensaje")) > 1 Then Producto = Producto & ", " & .fields("Mensaje")
                    Si_No = .fields("IVA")
                    If TipoFactura <> "FA" Then Si_No = False
                    Cta_Ventas = .fields("Cta_Ventas")
                    Real1 = CCur(TrimStrg(MidStrg(Rubros_Facturar(I), 76, 13)))   ' Valor
                   'MsgBox TrimStrg(MidStrg(Rubros_Facturar(I), 89, 13))
                    Real2 = CCur(TrimStrg(MidStrg(Rubros_Facturar(I), 89, 13)))   ' Descuento
                    Real5 = CCur(TrimStrg(MidStrg(Rubros_Facturar(I), 102, 13)))  ' Descuento PP
                    Real3 = Real1 - Real2 - Real5
                    Real4 = 0
                    If Si_No Then Real4 = CCur(Real3 * Porc_IVA)      'El valor del IVA
                   ' MsgBox MiMes & vbCrLf & NoMeses & vbCrLf & Codigo & vbCrLf & Codigo1 & vbCrLf _
                   '              & Real1 & vbCrLf & Real2 & vbCrLf & Real3 & vbCrLf & Real4
                    SetAdoAddNew "Asiento_F"
                    SetAdoFields "CODIGO", Codigo
                    SetAdoFields "CODIGO_L", CodigoL
                    SetAdoFields "PRODUCTO", Producto
                    SetAdoFields "CANT", 1
                    SetAdoFields "PRECIO", Real1
                    SetAdoFields "Total_Desc", Real2
                    SetAdoFields "Total_Desc2", Real5
                    SetAdoFields "TOTAL", Real1
                    SetAdoFields "Total_IVA", Real4
                    SetAdoFields "Cta", Cta_Ventas
                    SetAdoFields "Item", NumEmpresa
                    SetAdoFields "Codigo_Cliente", CodigoCliente
                   'SetAdoFields "RUTA", MidStrg("(" & Grupo_No & ") " & NombreCliente, 1, 50)
                    SetAdoFields "HABIT", Codigo2
                    SetAdoFields "Mes", MiMes
                    SetAdoFields "TICKET", Codigo1
                    SetAdoFields "CodigoU", CodigoUsuario
                    SetAdoFields "A_No", Contador
                    SetAdoUpdate
                    Contador = Contador + 1
                End If
               .MoveNext
             Loop
         End If
        End With
     End If
   Next I
   AdoAsientoF.Refresh
   Calculos_Totales_Factura FA
   Total_Desc = Redondear(FA.Descuento, 2)
   Total_Con_IVA = Redondear(FA.Con_IVA, 2)
   Total_Sin_IVA = Redondear(FA.Sin_IVA, 2)
   LabelDescuento.Caption = Format$(FA.Descuento, "#,##0.00")
   LabelDescuento2.Caption = Format$(FA.Descuento2, "#,##0.00")
   LabelSubTotal.Caption = Format$(FA.Sin_IVA, "#,##0.00")
   LabelConIVA.Caption = Format$(FA.Con_IVA, "#,##0.00")
   LabelIVA.Caption = Format$(FA.Total_IVA, "#,##0.00")
   LabelTotal.Caption = Format$(FA.Total_MN, "#,##0.00")
   TextCheque = Format$(FA.Total_MN, "#,##0.00")
   TextCheqNo = CGrupo.Text
End Sub

Private Sub MBHistorico_GotFocus()
   MarcarTexto MBHistorico
End Sub

Private Sub MBHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBHistorico_LostFocus()
   FechaValida MBHistorico
   MBHistorico = PrimerDiaMes(MBHistorico)
End Sub

Private Sub MBoxFecha_GotFocus()
   MarcarTexto MBoxFecha
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFecha_LostFocus()
   FechaValida MBoxFecha, True
   Validar_Porc_IVA MBoxFecha
   Label23.Caption = " Total Tarifa " & Format$(Porc_IVA * 100, "#0") & "%"
   FechaTexto1 = MBoxFecha.Text
   MBoxFechaV = CLongFecha(CFechaLong(MBoxFecha) + 15)
   Mifecha = BuscarFecha(MBoxFecha)
   FA.Fecha = MBoxFecha
   FA.Fecha_C = MBoxFecha
   FA.Fecha_V = MBoxFechaV
   FechaComp = FA.Fecha
   FA.Serie = "000000"
   FA.Autorizacion = Ninguno
   sSQL = "SELECT Codigo, Concepto, Serie, Autorizacion " _
        & "FROM Catalogo_Lineas " _
        & "WHERE TL <> " & Val(adFalse) & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fact = '" & FA.TC & "' " _
        & "AND Fecha <= #" & BuscarFecha(FA.Fecha) & "# " _
        & "AND Vencimiento >= #" & BuscarFecha(FA.Fecha) & "# " _
        & "ORDER BY Codigo "
   SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
   If AdoLinea.Recordset.RecordCount > 0 Then
      FA.Serie = AdoLinea.Recordset.fields("Serie")
      FA.Autorizacion = AdoLinea.Recordset.fields("Autorizacion")
   End If
   
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
   FechaValida MBoxFechaV
   FA.Fecha_V = MBoxFechaV
End Sub

Private Sub TextCI_GotFocus()
   MarcarTexto TextCI
End Sub

Private Sub TextCI_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

'Private Sub TextCI_KeyPress(KeyAscii As Integer)
'   KeyAscii = Solo_Letras_Numeros(KeyAscii)
'End Sub

Private Sub TextCI_LostFocus()
  DigVerif = Digito_Verificador(TextCI)
  If DigVerif = "-" Then
     Titulo = "FORMULARIO DE CONFIRMACION"
     Mensajes = "ADVERTENCIA: RUC/CEDULA INCORRECTA," & vbCrLf & vbCrLf _
              & "ESTE CODIGO ES UN PASAPORTE?"
     If BoxMensaje = vbYes Then
        Label18.Caption = Tipo_RUC_CI.Tipo_Beneficiario
        TxtDireccion.SetFocus
     Else
        TextCI.SetFocus
     End If
  Else
     Label18.Caption = Tipo_RUC_CI.Tipo_Beneficiario
     If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
        Label5.Caption = " RAZON SOCIAL"
     Else
        Label5.Caption = " PERSONA NATURAL"
     End If
     TxtDireccion.SetFocus
  End If
End Sub

Private Sub TextFacturaNo_GotFocus()
  MarcarTexto TextFacturaNo
End Sub

Private Sub TextFacturaNo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFacturaNo_LostFocus()
  If TextFacturaNo = "" Then TextFacturaNo = "0"
  If Val(TextFacturaNo) <= 0 Then TextFacturaNo = FA.Factura
End Sub

Private Sub TextBanco_GotFocus()
  MarcarTexto TextBanco
End Sub

Private Sub TextBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextBanco_LostFocus()
  TextoValido TextBanco, , True
End Sub

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
End Sub

Private Sub TextCheque_GotFocus()
  MarcarTexto TextCheque
End Sub

Private Sub TextCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCheque_LostFocus()
  TextoValido TextCheque, True
  Total_Bancos = Redondear(Val(CCur(TextCheque.Text)), 2)
  TotalCajaMN = FA.Total_MN - Total_Bancos - SubTotal_NC - Total_Anticipo
  SaldoDisp = FA.Total_MN - TotalCajaMN - Total_Bancos - Total_Anticipo - SubTotal_NC
  TxtEfectivo = Format$(TotalCajaMN, "#,##0.00")
  TextCheque.Text = Format$(Total_Bancos, "#,##0.00")
  LblCambio.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TextRepresentante_GotFocus()
  MarcarTexto TextRepresentante
End Sub

Private Sub TextRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRepresentante_LostFocus()
  TextoValido TextRepresentante, , True
End Sub

Private Sub Timer1_Timer()
  LblSaldo.Caption = "0.00"
  If Si_No_local And Total_Saldo_Pendiente > 0 Then
     LblSaldo.Caption = Format$(Total_Saldo_Pendiente, "#,##0.00")
  End If
  Si_No_local = Not (Si_No_local)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 'MsgBox Button.key
 RatonReloj
 Select Case Button.key
'''   Case "Matriculas"
'''        Actualizar_Matriculas
'''   Case "Buses"
'''        Actualizar_Buses
   Case "Deuda_Pendiente"
        Deuda_Pendiente
   Case "Recalcular_Saldos"
        FA.Fecha_Corte = FechaSistema
        FA.Fecha_Desde = FechaSistema
        FA.Fecha_Hasta = FechaSistema
        FA.TC = Ninguno
        FA.Serie = Ninguno
        FA.Factura = 0
        Actualizar_Abonos_Facturas_SP FA
        DCLinea.SetFocus
   Case "PreFacturas"
        FInsPreFacturas.Show 1
   Case "NuevoCliente"
        Nuevo = True
        NombreCliente = DCCliente.Text
        FacturasPension.Visible = False
        FClientesFlash.Show 1
        FacturasPension.Visible = True
        DCLinea.SetFocus
   Case "UpdateCliente"
        Actualiza_Datos_Cliente
   Case "LeerJS"
        URLHTTP = "C:\SISTEMA\JAVASCRIPT\estudiantes.html"
        XML = Replace(GetUrlSource(URLHTTP), """", "'")
        MsgBox XML
   Case "Salir"
        Unload FacturasPension
 End Select
 RatonNormal
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Resultado As Boolean
 'MsgBox ButtonMenu.key
  Select Case ButtonMenu.key
    Case "CarteraPDF"
         FechaInicial = MBHistorico.Text
         Reporte_Cartera_Clientes_SP FechaInicial, UltimoDiaMes(FechaSistema), TBeneficiario.Codigo
         Resultado = Reporte_Cartera_Clientes_PDF(FechaInicial, TBeneficiario.Codigo, False, True)
    Case "CarteraExcel"
         FechaInicial = MBHistorico.Text
         Reporte_Cartera_Clientes_SP FechaInicial, UltimoDiaMes(FechaSistema), TBeneficiario.Codigo
         Resultado = Reporte_Cartera_Clientes_PDF(FechaInicial, TBeneficiario.Codigo, True, False)
    Case "CarteraMail"
         FechaInicial = MBHistorico.Text
         Reporte_Cartera_Clientes_SP FechaInicial, UltimoDiaMes(FechaSistema), TBeneficiario.Codigo
         Resultado = Reporte_Cartera_Clientes_PDF(FechaInicial, TBeneficiario.Codigo, False, False)
         TMail.TipoDeEnvio = "CO"
         TMail.Asunto = "Estimado(a): " & VerPDF.NombreBeneficiario & ", usted tiene los siguientes pendientes."
         TMail.Mensaje = "Envio automatizado de su cartera pendiente." & vbCrLf _
                       & "NOTA: En caso de tener inconformidad con los valores detallados en su Estado de Cuenta, " _
                       & "comuniquese con atencion al Cliente."
         TMail.Adjunto = RutaDocumentoPDF
         
        'Enviamos lista de mails
         TMail.para = ""
         Insertar_Mail TMail.para, TBeneficiario.EmailR
         Insertar_Mail TMail.para, TBeneficiario.Email2
         Insertar_Mail TMail.para, TBeneficiario.Email1
         FEnviarCorreos.Show vbModal
  End Select
End Sub

Private Sub TxtCI_RUC_GotFocus()
  MarcarTexto TxtCI_RUC
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCI_RUC_LostFocus()
'''Dim pJSON As Object
'''    If Len(TxtCI_RUC) = 10 And Len(Token) > 1 Then
'''       TextoJSON = GetUrlSource(urlIdukay & TxtCI_RUC & RUC & NumEmpresa)
      'Limpia y pega texto en portapapeles
'''       Clipboard.Clear
'''       Clipboard.SetText urlIdukay & TxtCI_RUC & RUC & NumEmpresa
'''       Set pJSON = JSON.parse(TextoJSON)
'''       If Not (pJSON Is Nothing) Then
'''          If JSON.GetParserErrors <> "" Then
'''             MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
'''          Else
'''             Estudiante_DBF.cedula = pJSON.Item("user").Item("id_card")
'''             Estudiante_DBF.codest = Estudiante_DBF.cedula
'''             Estudiante_DBF.Fecha_Nac = pJSON.Item("user").Item("birthday")
'''             Estudiante_DBF.Sexo = pJSON.Item("user").Item("gender")
'''             Estudiante_DBF.Nombres = pJSON.Item("user").Item("surname") & " " & pJSON.Item("user").Item("second_surname") & " " _
'''                                    & pJSON.Item("user").Item("name") & " " & pJSON.Item("user").Item("second_name")
'''             Estudiante_DBF.NombreCurso = UCaseStrg(pJSON.Item("relational_data").Item("years").Item("grade").Item("show"))
'''             If Not IsNull(pJSON.Item("years").Item("group").Item("name")) Then Estudiante_DBF.Paralelo = pJSON.Item("years").Item("group").Item("name")
'''             Estudiante_DBF.pagador = pJSON.Item("relatives").Item("parent").Item("relational_data").Item("name").Item("show")
'''             Estudiante_DBF.cedular = pJSON.Item("relatives").Item("parent").Item("relational_data").Item("id_card").Item("show")
'''             Estudiante_DBF.emailpaga = pJSON.Item("relatives").Item("parent").Item("relational_data").Item("email")
'''             Estudiante_DBF.fonopaga = pJSON.Item("relatives").Item("parent").Item("phones").Item("mobile").Item("number")
'''             Estudiante_DBF.matriculado = pJSON.Item("relatives").Item("parent").Item("legal_person")
'''             Estudiante_DBF.direcpaga = pJSON.Item("relatives").Item("parent").Item("home_address")
'''           End If
'''       End If
'''    End If
End Sub

Private Sub TxtCodigoC_GotFocus()
  MarcarTexto TxtCodigoC
End Sub

Private Sub TxtDireccion_GotFocus()
  MarcarTexto TxtDireccion
End Sub

Private Sub TxtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDireccion_LostFocus()
  TextoValido TxtDireccion, , True
End Sub

Private Sub TxtDirS_GotFocus()
   MarcarTexto TxtDirS
End Sub

Private Sub TxtDirS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_GotFocus()
  MarcarTexto TxtEfectivo
End Sub

Private Sub TxtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_LostFocus()
  TextoValido TxtEfectivo, True, , 2
  TotalCajaMN = Redondear(Val(CCur(TxtEfectivo.Text)), 2)
  TxtEfectivo.Text = Format$(TotalCajaMN, "#,##0.00")
  SaldoDisp = FA.Total_MN - TotalCajaMN - Total_Bancos - SubTotal_NC - Total_Anticipo
  LblCambio.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TxtNC_GotFocus()
  MarcarTexto TxtNC
End Sub

Private Sub TxtNC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNC_LostFocus()
  TextoValido TxtNC, True
  SubTotal_NC = Redondear(Val(CCur(TxtNC.Text)), 2)
  TotalCajaMN = FA.Total_MN - Total_Bancos - SubTotal_NC - Total_Anticipo
  SaldoDisp = FA.Total_MN - TotalCajaMN - Total_Bancos - Total_Anticipo - SubTotal_NC
  TxtEfectivo = Format$(TotalCajaMN, "#,##0.00")
  TxtNC.Text = Format$(SubTotal_NC, "#,##0.00")
  LblCambio.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TxtEmail_GotFocus()
  MarcarTexto TxtEmail
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEmail_LostFocus()
  TxtEmail = LCase(TxtEmail)
  TextoValido TxtEmail
  
End Sub

Private Sub TxtDirS_LostFocus()
  TextoValido TxtDirS, , True
  If TxtDirS <> DireccionCli Then
     If Len(TrimStrg(TxtDirS)) <= 1 Then TxtDirS = DireccionCli
     DireccionCli = TxtDirS
  End If
End Sub

Private Sub TxtSaldoFavor_GotFocus()
  MarcarTexto TxtSaldoFavor
End Sub

Private Sub TxtSaldoFavor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtSaldoFavor_LostFocus()
  TextoValido TxtSaldoFavor, True
  Total_Anticipo = Redondear(Val(CCur(TxtSaldoFavor)), 2)
  TotalCajaMN = FA.Total_MN - Total_Bancos - SubTotal_NC - Total_Anticipo
  SaldoDisp = FA.Total_MN - TotalCajaMN - Total_Bancos - Total_Anticipo - SubTotal_NC
  TxtSaldoFavor = Format$(Total_Anticipo, "#,##0.00")
  TxtEfectivo = Format$(TotalCajaMN, "#,##0.00")
  LblCambio.Caption = Format$(SaldoDisp, "#,##0.00")
End Sub

Private Sub TxtTelefono_GotFocus()
  MarcarTexto TxtTelefono
End Sub

Private Sub TxtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtTelefono_LostFocus()
  TextoValido TxtTelefono, , True
  If TxtTelefono <> TelefCliente And Len(TelefCliente) > 1 Then
     If Len(TrimStrg(TxtTelefono)) <= 1 Then TxtTelefono = TelefCliente
     TelefCliente = TxtTelefono
  End If
End Sub

Private Sub TextInteres_GotFocus()
  MarcarTexto TextInteres
End Sub

Private Sub TextInteres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextInteres_LostFocus()
  If MidStrg(TextInteres, Len(TextInteres), 1) = "%" Then
     Valor = MidStrg(TextInteres, 1, Len(TextInteres) - 1)
     TextInteres = Valor * Val(LabelTotal.Caption) / 100
  End If
  TextoValido TextInteres, True
End Sub

'''Public Sub Actualizar_Buses()
'''  FechaInicial = Dato_DBF.FechaI
'''  FechaFinal = Dato_DBF.FechaF
'''  If IsDate(FechaInicial) And IsDate(FechaFinal) And Actualiza_Buses Then
'''     Progreso_Barra.Mensaje_Box = "Actualizar Pagos de Buses"
'''     Progreso_Iniciar
'''
'''     sSQL = "SELECT Codigo, Cliente, CI_RUC, Grupo, Plan_Afiliado " _
'''          & "FROM Clientes " _
'''          & "WHERE MidStrg(Plan_Afiliado, 1, 3) = 'BUS' " _
'''          & "AND LEN(Plan_Afiliado) = 6 " _
'''          & "AND MidStrg(Grupo,1,3) <> 'RET' " _
'''          & "ORDER BY Plan_Afiliado,Cliente "
'''     Select_Adodc AdoAux, sSQL
'''     With AdoAux.Recordset
'''      If .RecordCount > 0 Then
'''          Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo & .RecordCount
'''          Do While Not .EOF
'''             RatonReloj
'''            'MsgBox .Fields("Cliente") & vbCrLf & .Fields("CI_RUC")
'''             CodigoA = .fields("Codigo")
'''             CodigoB = .fields("Plan_Afiliado")
'''             CodigoC = .fields("Grupo")
'''             Progreso_Barra.Mensaje_Box = "Actualizando: " & .fields("Cliente") & ", " & CodigoB
'''             Progreso_Esperar
'''             Actualizar_Bus CodigoA, CodigoB, CodigoC
'''             RatonNormal
'''            .MoveNext
'''          Loop
'''      End If
'''     End With
'''     Progreso_Final
'''     MsgBox "Proceso Terminado"
'''  Else
'''     RatonNormal
'''     MsgBox "No se realizo este proceso"
'''  End If
'''End Sub

Public Sub Deuda_Pendiente()
Dim Ini_X As Single
Dim Ini_Y As Single
Dim tipoDeLetra As String
    
    NombreCliente = DCCliente
    tipoDeLetra = TipoCourier
   'Geneeramos el documento
    tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = "Deuda_de_" & Replace(DCCliente, " ", "_")
    tPrint.TituloArchivo = "Deuda_de_" & Replace(DCCliente, " ", "_")
    tPrint.TipoLetra = tipoDeLetra
    tPrint.OrientacionPagina = 1
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = True
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
    Ini_X = 1
    cPrint.letraTipo tipoDeLetra, 9
    cPrint.printCuadro 1, 1, 1, 1, Negro, "B", 0.1
    cPrint.printImagen LogoTipo, Ini_X + 0.2, 1, 3, 1.2
    PosLinea = 0.6
    cPrint.printTexto Ini_X + 4, PosLinea, UCaseStrg(RazonSocial)
    PosLinea = PosLinea + 0.35
    If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
       cPrint.printTexto Ini_X + 4, PosLinea, UCaseStrg(NombreComercial)
       PosLinea = PosLinea + 0.35
    End If
    cPrint.printTexto Ini_X + 4, PosLinea, "R.U.C.: " & RUC
    PosLinea = PosLinea + 0.35
    cPrint.letraTipo tipoDeLetra, 5
    cPrint.printTexto Ini_X + 4, PosLinea, "Dirección: " & Direccion
    PosLinea = PosLinea + 0.25
    If UCaseStrg(Direccion) <> UCaseStrg(DireccionEstab) Then
       cPrint.printTexto Ini_X + 4, PosLinea, "Sucursal: " & DireccionEstab
       PosLinea = PosLinea + 0.25
    End If
    cPrint.printTexto Ini_X + 4, PosLinea, "Telefono(s): " & Telefono1 & " / " & Telefono2 & " / " & FAX
    PosLinea = PosLinea + 0.25
    cPrint.printTexto Ini_X + 4, PosLinea, ULCase(NombreCiudad) & " - Ecuador"
    PosLinea = 2.6
    cPrint.letraTipo tipoDeLetra, 8
    cPrint.printTexto 1.5, PosLinea, "DEUDA PENDIENTE DE"
    'cPrint.printTexto 14.5, PosLinea, "CODIGO DEL BANCO"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.5, PosLinea, "CURSO DEL ESTUDIANTE"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.5, PosLinea, "REPRESENTANTE LEGAL"
    cPrint.printTexto 14.5, PosLinea, "CEDULA DE IDENTIDAD"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 1.5, PosLinea, "CORREO ELECTRONICO"
    cPrint.printTexto 14.5, PosLinea, "TELEFONO DE CONTACTO"
    PosLinea = PosLinea + 0.4
    PosLinea = 2.6
    cPrint.printTexto 5, PosLinea, ":" & NombreCliente
    'cPrint.printTexto 18, PosLinea, ":" & CodigoB
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 5, PosLinea, ":" & TxtDirS
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 5, PosLinea, ":" & TextRepresentante
    cPrint.printTexto 18, PosLinea, ":" & TextCI
    PosLinea = PosLinea + 0.4
    cPrint.printTexto 5, PosLinea, ":" & TxtEmail
    cPrint.printTexto 18, PosLinea, ":" & TxtTelefono
    PosLinea = PosLinea + 0.4
''    TxtDireccion
''    Label20.Caption
    Ini_Y = PosLinea
    cPrint.letraTipo tipoDeLetra, 7
    cPrint.printTexto 1.5, PosLinea, TrimStrg(Label11.Caption)
    PosLinea = PosLinea + 0.4
    Valor = 0
    For I = 0 To LstMeses.ListCount - 1
     If LstMeses.Selected(I) Then
        cPrint.printTexto 1.5, PosLinea, LstMeses.List(I)
        Valor = Valor + Val(TrimStrg(MidStrg(LstMeses.List(I), 115, 13)))
        PosLinea = PosLinea + 0.35
     End If
    Next I
    PosLinea = PosLinea + 0.1
    cPrint.printCuadro 1.4, Ini_Y - 0.05, 19.9, PosLinea - 0.7, Negro, "B", 0.1
    cPrint.printLinea 1.4, Ini_Y + 0.3, 19.8, Ini_Y + 0.3, Negro, 0.1
    cPrint.printTexto Ini_X + 0.5, PosLinea, "CORTE AL " & FechaStrg(FechaSistema)
    cPrint.printTexto Ini_X + 14.5, PosLinea, "TOTAL A PAGAR  USD "
    cPrint.printVariable Ini_X + 16.65, PosLinea, Valor, , Rojo
    PosLinea = PosLinea + 0.5
    Cadena = "Los datos presentados en este reporte, reflejan los valores pendientes de pago, por concepto de Pensiones Educativas."
    PosLinea = cPrint.printTextoMultiple(Ini_X + 0.5, PosLinea, Cadena, 18)
    'PosLinea = PosLinea + 0.5
    'Cadena = "Este documento caduca en 5 días hábiles."
    'PosLinea = cPrint.printTextoMultiple(Ini_X + 0.5, PosLinea, Cadena, 18)
    PosLinea = PosLinea + 2.5
    cPrint.printTexto Ini_X + 0.4, PosLinea, String(Len(NombreUsuario), "_") & "_"
    PosLinea = PosLinea + 0.4
    cPrint.printTexto Ini_X + 0.5, PosLinea, NombreUsuario
    PosLinea = PosLinea + 0.4
    cPrint.printTexto Ini_X + 0.5, PosLinea, "COLECTURIA"
    cPrint.finalizaImpresion
    RatonNormal
    Titulo = "FORMULARIO DE CONFIRMACION"
    Mensajes = "ENVIAR REPORTE POR MAIL:" & vbCrLf & vbCrLf _
             & "CORREO(S) DE ENVIO: " & TMail.para
    If BoxMensaje = vbYes Then
       TMail.TipoDeEnvio = "CO"
       TMail.Asunto = "Estimado(a): " & NombreCliente & ", usted tiene los siguientes pendientes."
       TMail.Mensaje = "Envio automatizado de su cartera pendiente." & vbCrLf _
                     & "NOTA: En caso de tener inconformidad con los valores detallados en su Estado de Cuenta, comuniquese con atencion al Cliente."
       TMail.Adjunto = RutaDocumentoPDF
       FEnviarCorreos.Show vbModal
    End If
End Sub

Public Function Saldo_De_Anticipos(CtaAnticipos As String) As Currency
Dim TotAntipos As Currency
Dim AuxCtaAnticipo As String
  TotAntipos = 0
  AuxCtaAnticipo = CtaAnticipos
  If Len(AuxCtaAnticipo) <= 1 Then AuxCtaAnticipo = "0"
  sSQL = "SELECT SUM(Creditos-Debitos) As Saldo_Pendiente " _
       & "FROM Trans_SubCtas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Codigo = '" & FA.CodigoC & "' " _
       & "AND Cta = '" & AuxCtaAnticipo & "' " _
       & "AND T = 'N' "
  Select_AdoDB AdoDBTemp, sSQL
  If AdoDBTemp.RecordCount > 0 And Not IsNull(AdoDBTemp.fields("Saldo_Pendiente")) Then TotAntipos = AdoDBTemp.fields("Saldo_Pendiente")
  AdoDBTemp.Close
  sSQL = "SELECT SUM(Abono) As Abonado " _
       & "FROM Trans_Abonos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND CodigoC = '" & FA.CodigoC & "' " _
       & "AND Cta = '" & AuxCtaAnticipo & "' " _
       & "AND C = " & Val(adFalse) & " "
  Select_AdoDB AdoDBTemp, sSQL
  If AdoDBTemp.RecordCount > 0 And Not IsNull(AdoDBTemp.fields("Abonado")) Then TotAntipos = TotAntipos - AdoDBTemp.fields("Abonado")
  AdoDBTemp.Close
 'MsgBox TotAntipos
  Saldo_De_Anticipos = TotAntipos
End Function
