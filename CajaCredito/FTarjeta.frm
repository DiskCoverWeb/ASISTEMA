VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form FTarjeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   7605
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5880
      Top             =   3255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   3675
   End
   Begin VB.Frame Frame1 
      Caption         =   "INFORMACION DE TRANSMISION:"
      Height          =   2745
      Left            =   105
      TabIndex        =   18
      Top             =   4200
      Width           =   7365
      Begin VB.TextBox TextEstado 
         Height          =   1485
         Left            =   3675
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   1155
         Width           =   3585
      End
      Begin VB.TextBox TextTraza 
         Height          =   1485
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   1155
         Width           =   3585
      End
      Begin VB.TextBox txtPaquetesSend 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   210
         Width           =   540
      End
      Begin VB.TextBox txtPaquetesGet 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   525
         Width           =   540
      End
      Begin VB.TextBox TxtRemotePort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4515
         TabIndex        =   28
         Text            =   "1002"
         Top             =   525
         Width           =   750
      End
      Begin VB.TextBox txtLocalPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4515
         TabIndex        =   27
         Text            =   "1001"
         Top             =   210
         Width           =   750
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Recibir"
         Height          =   645
         Left            =   5985
         Picture         =   "FTarjeta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Enviar"
         Height          =   645
         Left            =   5355
         Picture         =   "FTarjeta.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   210
         Width           =   645
      End
      Begin VB.TextBox TxtRemoteHost 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   19
         Text            =   "192.168.0.2"
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Puerto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3465
         TabIndex        =   30
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Remote Host"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   20
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Puerto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3465
         TabIndex        =   29
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ESTADO DE LA TRAMA"
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
         TabIndex        =   24
         Top             =   945
         Width           =   3585
      End
      Begin VB.Label LabelLocalIP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "192.168.0.100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Top             =   210
         Width           =   1800
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRAMA DE ENVIO/RECEPCION"
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
         TabIndex        =   23
         Top             =   945
         Width           =   3585
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   22
         Top             =   210
         Width           =   1590
      End
   End
   Begin MSDataListLib.DataList DLTP 
      Bindings        =   "FTarjeta.frx":0C30
      DataSource      =   "AdoTP"
      Height          =   2010
      Left            =   105
      TabIndex        =   17
      Top             =   2100
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3545
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   105
      Top             =   3255
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
      Caption         =   "Asientos"
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   105
      Top             =   2310
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
      Caption         =   "TP"
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
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5670
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "FTarjeta.frx":0C44
      Top             =   2100
      Width           =   1695
   End
   Begin VB.TextBox TextLinea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5670
      MaxLength       =   8
      TabIndex        =   12
      Top             =   2835
      Width           =   1380
   End
   Begin VB.TextBox TxtCheque 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5670
      MaxLength       =   8
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1380
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
      Height          =   855
      Left            =   3990
      Picture         =   "FTarjeta.frx":0C4B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3255
      Width           =   855
   End
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
      Height          =   855
      Left            =   4935
      Picture         =   "FTarjeta.frx":108D
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3255
      Width           =   855
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCC-CCCCCCCCC-C"
      Mask            =   "#######-########-#"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   105
      Top             =   2625
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
      Caption         =   "CtaNo"
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
      Left            =   105
      Top             =   2940
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
   Begin VB.PictureBox CmdEstado 
      Height          =   750
      Left            =   6825
      ScaleHeight     =   690
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   105
      Width           =   540
   End
   Begin VB.Label LabelCta_No 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000000"
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
      TabIndex        =   33
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado de Cta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2940
      TabIndex        =   34
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label LabelEstado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NORMAL"
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
      Left            =   4830
      TabIndex        =   15
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado de Cta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4830
      TabIndex        =   16
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo de Transaccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1785
      Width           =   3795
   End
   Begin VB.Label LabelSocio 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Width           =   7260
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Monto Trans."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      TabIndex        =   7
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Label LabelDisponible 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5670
      TabIndex        =   5
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Disponible"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      TabIndex        =   4
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre y Apellidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   945
      Width           =   7260
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea No."
      BeginProperty Font 
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
      TabIndex        =   11
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tarjeta de Debito No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2745
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cheque No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3990
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "FTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function DimeEstado() As String
    Select Case Winsock1.State
    Case 0
        DimeEstado = "Cerrado"
    Case 1
        DimeEstado = "Abierto"
    Case 2
        DimeEstado = "Escuchando"
    Case 3
        DimeEstado = "Conexión pendiente..."
    Case 4
        DimeEstado = "Resolviendo host..."
    Case 5
        DimeEstado = "Host resuelto"
    Case 6
        DimeEstado = "Conectando..."
    Case 7
        DimeEstado = "Conectado"
    Case 8
        DimeEstado = "Cerrando..."
    Case 9
        DimeEstado = "ERROR !!"
    Case Else
        DimeEstado = "Desconocido (" & Winsock1.State & ")"
    End Select
End Function

Private Sub EscribeEnTraza(Optional ByVal strMSG As String)
    Dim strTraza As String
    With TextEstado
        strTraza = Format(Now, "hh:nn:ss") & _
            " --> (Estado del socket: " & DimeEstado & ")" & vbNewLine
        If strMSG <> "" Then strTraza = strTraza & "   " & strMSG & vbNewLine
        .Text = .Text & strTraza
        .SelStart = Len(.Text) - 1
        If Len(.Text) > 65000 Then .Text = ""
    End With
End Sub

Public Sub InsertarMontosTarjeta(DtaCta As Adodc, _
                                 DtaBanc As Adodc, _
                                 CuentaNo As MaskEdBox, _
                                 TDebe As Currency, _
                                 THaber As Currency, _
                                 NoCheque As String, _
                                 NomBanco As String)
  If CuentaNo.Text <> "0000000-00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
 'Insertar Transacciones de Libreta
  sSQL = "SELECT TOP 1 * FROM Trans_Tarjetas " _
       & "WHERE Tarjeta_No = '" & CuentaNo.Text & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           ID_Trans = .Fields("IDT")
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Tarjeta_No") = CuentaNo.Text
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
      .Fields("T") = Normal
      .Fields("CodigoU") = CodigoUsuario
       If NumeroLineas >= 255 Then NumeroLineas = 254
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = Moneda_US
      .Fields("Ticket") = NoCheque
      .Update
  End With
  End If
End Sub

Public Sub ListarTarjetas(CuentaNo As MaskEdBox)
  Mi_Cta = False
  SaldoDisp = 0:  SaldoCont = 0
  TotalEncaje = 0: NumeroLineas = 0
  Moneda_US = False: LabelSocio.Caption = ""
  T = Normal: TextLinea.Text = NumeroLineas
  LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
  sSQL = "SELECT C.*,TT " _
       & "FROM Tarjetas As T,Clientes_Datos_Extras As C " _
       & "WHERE Tarjeta_No = '" & CuentaNo.Text & "' " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND T.Cuenta_No = C.Cuenta_No "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
       Mi_Cta = True
       ConLibreta = True
       T = .Fields("TT")
       LabelCta_No.Caption = .Fields("Cuenta_No")
       LabelSocio.Caption = .Fields("Nombres") & "  " & .Fields("Apellidos")
       sSQL = "SELECT TOP 1 * " _
            & "FROM Trans_Tarjetas " _
            & "WHERE Tarjeta_No = '" & CuentaNo.Text & "' " _
            & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
       SelectAdodc AdoCtaNo, sSQL
       If AdoCtaNo.Recordset.RecordCount > 0 Then
          NumeroLineas = AdoCtaNo.Recordset.Fields("ID")
          SaldoDisp = AdoCtaNo.Recordset.Fields("Saldo_Disp")
       End If
       TextLinea.Text = NumeroLineas
       LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
       If T = Procesado Then
          LabelEstado.Caption = "Apertura"
          sSQL = "SELECT Proceso " _
               & "FROM Catalogo_Proceso " _
               & "WHERE MidStrg(TP,1,4) = 'APET' " _
               & "AND Nivel = 4 " _
               & "ORDER BY Proceso "
       ElseIf T = "B" Then
          LabelEstado.Caption = "Bloqueada"
          MsgBox "Tarjeta Bloqueada, no podra realizar transacciones"
       Else
          LabelEstado.Caption = "Normal"
          sSQL = "SELECT Proceso " _
               & "FROM Catalogo_Proceso " _
               & "WHERE MidStrg(TP,1,4) <> 'APET' " _
               & "AND Nivel = 4 " _
               & "ORDER BY Proceso "
       End If
       If T = "B" Then
          MBoxCuenta.SetFocus
       Else
          SelectDBList DLTP, AdoTP, sSQL, "Proceso"
          DLTP.SetFocus
       End If
   Else
       ConLibreta = False
       MsgBox "Esta cuenta no exite o esta anulada "
       LabelEstado.Caption = "NINGUNO"
       LabelSocio.Caption = "TRANSACCION SIN LIBRETA"
       MBoxCuenta.SetFocus
   End If
  End With
End Sub

Private Sub Command1_Click()
  If CDbl(TxtMonto.Text) > 0 Then
  Monto_Total = 0
  TextoProc = DLTP.Text
  TextoValidoVar TextoProc
  sSQL = "SELECT * FROM Catalogo_Proceso " _
       & "WHERE Proceso = '" & TextoProc & "' "
  SelectAdodc AdoCtaNo, sSQL
  With AdoCtaNo.Recordset
   If .RecordCount > 0 Then
       NivelNo = .Fields("Cmds")
       NivelCta = .Fields("Nivel")
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM MontoA_Tarjeta "
  SelectAdodc AdoCtaNo, sSQL
  If AdoCtaNo.Recordset.RecordCount > 0 Then
     MontoAper = AdoCtaNo.Recordset.Fields("Monto_Aper")
  End If
  If NumCheque = "" Then NumCheque = Ninguno
  If NombreBanco = "" Then NombreBanco = Ninguno
  TextoTraza = NivelNo
  TextoTraza = TextoTraza & MBoxCuenta.Text
  TextoTraza = TextoTraza & "01100"
  TextoTraza = TextoTraza & Format(CDbl(TxtMonto.Text), "0000000000.00")
  TextoTraza = TextoTraza & Format(FechaSistema, "YYYYMMDD")
  TextoTraza = TextoTraza & Format(FechaSistema, "MMDD")
  TextoTraza = TextoTraza & Format(NumeroLineas, "000000")
  TextoTraza = TextoTraza & String(12, "0")
  TextoTraza = TextoTraza & "593"
  TextoTraza = TextoTraza & String(12, "0")
  TextoTraza = TextoTraza & String(11, "0")
  TextoTraza = TextoTraza & "60194107"
  TextoTraza = TextoTraza & String(12, "0")
  TextoTraza = TextoTraza & String(3, "0")
  TextoTraza = TextoTraza & String(8, "0")
  TextoTraza = TextoTraza & String(15, "0")
  TextoTraza = TextoTraza & String(42, "0")
  TextoTraza = TextoTraza & "840"
 'Datos Transaccion
  TextoTraza = TextoTraza & "10"
  Select Case DC_Caja
    Case "C"
         InsertarMontosTarjeta AdoCuentas, AdoCtaNo, MBoxCuenta, 0, CDbl(TxtMonto.Text), NumCheque, NombreBanco
         If T = Procesado Then
            NumeroLineas = NumeroLineas + 1
            TipoProc = "N/DG"
            InsertarMontosTarjeta AdoCuentas, AdoCtaNo, MBoxCuenta, MontoAper, 0, NumCheque, NombreBanco
         End If
         sSQL = "UPDATE Tarjetas " _
              & "SET TT = '" & Normal & "' " _
              & "WHERE Tarjeta_No = '" & MBoxCuenta.Text & "' "
         ConectarAdoExecute sSQL
         TextoTraza = TextoTraza & "01840C"
    Case "D"
         InsertarMontosTarjeta AdoCuentas, AdoCtaNo, MBoxCuenta, CDbl(TxtMonto.Text), 0, NumCheque, NombreBanco
         TextoTraza = TextoTraza & "02840D"
  End Select
  End If
  TextoTraza = TextoTraza & Format(CDbl(LabelDisponible.Caption), "0000000000.00")
  TextoTraza = TextoTraza & Format(CDbl(LabelDisponible.Caption), "0000000000.00")
  TextoTraza = TextoTraza & String(120 - 32, "0")
  TextoTraza = TextoTraza & String(310, "0")
  TextoTraza = TextoTraza & LabelCta_No.Caption
  TextoTraza = TextoTraza & String(12, "0")
  TextTraza.Text = TextoTraza
  MBoxCuenta.SetFocus
End Sub

Private Sub Command2_Click()
   EscribeEnTraza "<FIN>"
  'Si el socket no está cerrado, lo cierra
   If Winsock1.State <> 0 Then Winsock1.Close
   Unload FTarjeta
End Sub

Private Sub Command3_Click()
 TextoTraza = TextTraza.Text
'Manda el MSG por defecto si la conexión está abierta
 With Winsock1
   If .State = 7 Then
       EscribeEnTraza "<Mandando paquete...>"
      .SendData TextoTraza
   Else
      .Close
      .RemoteHost = VRemoteHost
      .RemotePort = VRemotePort
      .Connect
   End If
 End With
End Sub

Private Sub cmdEscucha_Click()
    
    'Si el socket está cerrado, se pone en escucha
'''    With Me.TCPSocket
'''        If .State = "0" Then
'''            EscribeEnTraza "<Escucha>"
'''            'Apariencia form
'''            Me.txtPaquetesGet.Text = "0"
'''            Me.txtPaquetesSend.Text = "0"
'''            .LocalPort = Me.txtLocalPort.Text
'''            .Listen
'''        End If
'''    End With
    
End Sub

Private Sub Command4_Click()

  'Si el socket está cerrado, se pone en escucha
   With Winsock1
     If .State <> "0" Then
         EscribeEnTraza "<Escucha>"
        .LocalPort = VLocalPort
        .Listen
     End If
    'Sólo puede conectar si el socket está cerrado o escuchando...
     EscribeEnTraza "<Conecta>"
    'Apariencia form
     txtPaquetesGet.Text = "0"
     txtPaquetesSend.Text = "0"
    'Establece las propiedades del socket y conecta
    .RemoteHost = VRemoteHost
    .RemotePort = VRemotePort
    .Connect
    'If .State = 7 Then .SendData TextoTraza
   End With
End Sub

Private Sub DLTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLTP_LostFocus()
  Cadena = DLTP.Text
  If Cadena = "" Then Cadena = Ninguno
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE Proceso = '" & Cadena & "' "
  SelectAdodc AdoCtaNo, sSQL
  With AdoCtaNo.Recordset
   If .RecordCount > 0 Then
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       NivelNo = .Fields("Cmds")
       TipoGrupo = .Fields("Cheque")
       Label10.Visible = TipoGrupo
       TxtCheque.Visible = TipoGrupo
   End If
  End With
End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  'If Supervisor = False Then
  '   Command1.Enabled = CNivel(3) Or CNivel(4) Or CNivel(6)
  '   Command3.Enabled = CNivel(3) Or CNivel(4) Or CNivel(6)
  'End If
  sSQL = "SELECT Proceso " _
       & "FROM Catalogo_Proceso " _
       & "WHERE Nivel = 4 " _
       & "ORDER BY Proceso "
  SelectDBList DLTP, AdoTP, sSQL, "Proceso"
  FTarjeta.Caption = "OPERACIONES DE TARJETAS DE DEBITOS "
  TextoTraza = "SOCKET_HOLA"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FTarjeta
  ConectarAdodc AdoTP
  ConectarAdodc AdoCtaNo
  ConectarAdodc AdoCuentas
  ConectarAdodc AdoAsientos
  
  VLocalHost = "192.168.1.125"
  VLocalHostName = Winsock1.LocalHostName
  VLocalPort = 1001
  VRemoteHost = "192.168.1.179"
  VRemotePort = 5102
  TxtRemoteHost.Text = VRemoteHost
  TxtRemotePort.Text = VRemotePort
  LabelLocalIP.Caption = VLocalHost
  
  Winsock1.RemoteHost = VRemoteHost
  Winsock1.RemotePort = VRemotePort
  Winsock1.Connect
  Timer1.Interval = 500
End Sub

Private Sub MBoxCuenta_GotFocus()
  MarcarTexto MBoxCuenta
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCuenta_LostFocus()
  NombreBanco = ""
  NumCheque = ""
  If MBoxCuenta.Text = "0000000-00000000-0" Then
     MsgBox "Esta cuenta no esta permitida"
     MBoxCuenta.SetFocus
  Else
     Mi_Cta = False
     ListarTarjetas MBoxCuenta
  End If
End Sub

Private Sub TextLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Timer1_Timer()
Static intEstadoAnterior As Integer
 If Winsock1.State = "8" Then
    Winsock1.Close
    Command4_Click   'Escucha
 End If
 
 If Winsock1.State = "0" Then
    EscribeEnTraza "<Escuchando > "
    'Winsock1.LocalIP = VLocalHost
    Winsock1.LocalPort = VLocalPort
    Winsock1.Listen
 End If
  'Muestra el estado en pantalla, si ha cambiado desde la última vez.
'  If Winsock1.State <> 7 Then
'     Winsock1.Close
'     Winsock1.Connect
'     Winsock1.RemoteHost = VRemoteHost
'     Winsock1.RemotePort = VRemotePort

'   End If
'
'   Select Case Winsock1.State
'
'     Case 0, 9: CmdEstado.Picture = LoadPicture(RutaSistema & "\ICONOS\No.ico")
'             Winsock1.Close
'             For I = 1 To 1000: Next I
'             Winsock1.RemoteHost = VRemoteHost
'             Winsock1.RemotePort = VRemotePort
'             Winsock1.Connect
'     Case 6: CmdEstado.Picture = LoadPicture(RutaSistema & "\ICONOS\Rojo.ico")
'     Case Else
'          CmdEstado.Picture = LoadPicture(RutaSistema & "\ICONOS\Visto.ico")
'   End Select
End Sub

Private Sub TxtCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCheque_LostFocus()
  TextoValido TxtCheque, , True
  NumCheque = TxtCheque.Text
  NoCheque = TxtCheque.Text
End Sub

Private Sub TxtMonto_GotFocus()
  TxtMonto.Text = ""
End Sub

Private Sub TxtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMonto_LostFocus()
  TextoValido TxtMonto, True
  TxtMonto.Text = Format(TxtMonto.Text, "#,##0.00")
  If SaldoDisp > 0 And ConLibreta Then
    If (DC_Caja = "D") And (CDbl(TxtMonto.Text) > CDbl(LabelDisponible.Caption)) Then
       MsgBox "Excede el saldo disponible"
       TxtMonto.SetFocus
    End If
  End If
End Sub

Private Sub Winsock1_Close()
  CmdEstado.Picture = LoadPicture(RutaSistema & "\ICONOS\No.ico")
  EscribeEnTraza "<El Host ha cerrado el socket>"
End Sub

Private Sub Winsock1_Connect()
 CmdEstado.Picture = LoadPicture(RutaSistema & "\ICONOS\Visto.ico")
 EscribeEnTraza "<Se ha conectado al Host>"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'Si el socket está cerrado acepta la petición
    MsgBox Winsock1.State
    With Winsock1
     If .State = 2 Then
         EscribeEnTraza "<Acepta la petición " & requestID & ">"
        .Close
        .Accept requestID
     End If
    End With
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
    Winsock1.GetData strdata
    If strdata = TextoTraza Then
        EscribeEnTraza "<Se ha recibido un MSG correctamente>"
        txtPaquetesGet.Text = CLng(txtPaquetesGet.Text) + 1
     Else
        EscribeEnTraza "<Se ha recibido un MSG no válido:" & vbNewLine _
                       & "MSG: " & strdata & ">"
        TextTraza.Text = strdata
    End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, _
                            Description As String, _
                            ByVal Scode As Long, _
                            ByVal Source As String, _
                            ByVal HelpFile As String, _
                            ByVal HelpContext As Long, _
                            CancelDisplay As Boolean)
                            
    'Cancela la muestra del msg por defecto
    CancelDisplay = True
    'Muestra traza
    EscribeEnTraza "<Error en el socket:" & vbNewLine & _
        "   * Número: " & Number & vbNewLine & _
        "   * Descripción: " & Description & vbNewLine & _
        "   * Código: " & Scode & vbNewLine & _
        "   * Origen: " & Source & ">" & vbNewLine
    'Desconecta
    If Winsock1.State <> 0 Then
      'Muestra traza
       EscribeEnTraza "<Desconecta>"
       Winsock1.Close
       'MsgBox "Seccion terminada"
    End If
End Sub

Private Sub Winsock1_SendComplete()
    EscribeEnTraza "<Envío correcto>"
    txtPaquetesSend.Text = CLng(txtPaquetesSend.Text) + 1
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, _
                                   ByVal bytesRemaining As Long)
    EscribeEnTraza "<Enviando: quedan " & bytesRemaining & " bytes>"
End Sub

