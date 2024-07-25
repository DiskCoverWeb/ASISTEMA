VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AbonoRetencion 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RETENCIONES EN LA FUENTE Y DEL I.V.A."
   ClientHeight    =   4680
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextBanco 
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
      Left            =   5670
      MaxLength       =   49
      TabIndex        =   19
      Text            =   "0000000000"
      Top             =   1260
      Width           =   3060
   End
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   5985
      Top             =   4620
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "RetIvaBienes"
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
   Begin VB.CheckBox ChRetF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Retención Fuente"
      Height          =   225
      Left            =   210
      TabIndex        =   34
      Top             =   3780
      Width           =   1905
   End
   Begin VB.Frame FrmRet 
      BackColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   105
      TabIndex        =   35
      Top             =   3570
      Visible         =   0   'False
      Width           =   8625
      Begin VB.TextBox TextRet 
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
         Height          =   360
         Left            =   6825
         MaxLength       =   10
         TabIndex        =   42
         Text            =   "0"
         Top             =   525
         Width           =   1695
      End
      Begin VB.TextBox TextPorc 
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
         Height          =   360
         Left            =   6090
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "0"
         Top             =   525
         Width           =   750
      End
      Begin MSDataListLib.DataCombo DCRetFuente 
         Bindings        =   "AbonoRet.frx":0000
         DataSource      =   "AdoRetFuente"
         Height          =   315
         Left            =   105
         TabIndex        =   36
         Top             =   525
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSDataListLib.DataCombo DCCodRet 
         Bindings        =   "AbonoRet.frx":001B
         DataSource      =   "AdoCodRet"
         Height          =   315
         Left            =   5145
         TabIndex        =   38
         ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
         Top             =   525
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Valor Retenido"
         Height          =   225
         Left            =   6825
         TabIndex        =   41
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "%"
         Height          =   225
         Left            =   6090
         TabIndex        =   39
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Codigo"
         Height          =   225
         Left            =   5250
         TabIndex        =   37
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.CheckBox ChRetS 
      BackColor       =   &H00C0FFFF&
      Caption         =   "IVA en Servicios"
      Height          =   225
      Left            =   210
      TabIndex        =   27
      Top             =   2835
      Width           =   1800
   End
   Begin VB.Frame FrmIVAS 
      BackColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   105
      TabIndex        =   28
      Top             =   2625
      Visible         =   0   'False
      Width           =   8625
      Begin VB.ComboBox CServicio 
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
         Left            =   6090
         TabIndex        =   31
         Text            =   "100"
         Top             =   525
         Width           =   750
      End
      Begin VB.TextBox TextRetIVAS 
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
         Left            =   6825
         MaxLength       =   10
         TabIndex        =   33
         Text            =   "0"
         Top             =   525
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "AbonoRet.frx":0033
         DataSource      =   "AdoRetIvaServicio"
         Height          =   315
         Left            =   105
         TabIndex        =   29
         Top             =   525
         Width           =   6000
         _ExtentX        =   10583
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
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Valor Retenido"
         Height          =   225
         Left            =   6825
         TabIndex        =   32
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "%"
         Height          =   225
         Left            =   6090
         TabIndex        =   30
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CheckBox ChRetB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "IVA en Bienes"
      Height          =   225
      Left            =   210
      TabIndex        =   20
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Frame FrmIVAB 
      BackColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   105
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   8625
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "AbonoRet.frx":0053
         DataSource      =   "AdoRetIvaBienes"
         Height          =   315
         Left            =   105
         TabIndex        =   22
         Top             =   525
         Width           =   6000
         _ExtentX        =   10583
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
      Begin VB.TextBox TextRetIVAB 
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
         Left            =   6825
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "0"
         Top             =   525
         Width           =   1695
      End
      Begin VB.ComboBox CBienes 
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
         Left            =   6090
         TabIndex        =   24
         Text            =   "100"
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Valor Retenido"
         Height          =   225
         Left            =   6825
         TabIndex        =   25
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "%"
         Height          =   225
         Left            =   6090
         TabIndex        =   23
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.TextBox TextCheqNo 
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
      Left            =   4830
      MaxLength       =   8
      TabIndex        =   17
      Text            =   "001001"
      Top             =   1260
      Width           =   855
   End
   Begin VB.TextBox TxtRecibo 
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
      Height          =   330
      Left            =   1575
      MaxLength       =   14
      TabIndex        =   1
      Text            =   "0"
      Top             =   105
      Width           =   1275
   End
   Begin VB.CheckBox CheqRecibo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&RECIBO No."
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   1380
   End
   Begin VB.TextBox TextCompRet 
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
      Left            =   3465
      MaxLength       =   9
      TabIndex        =   15
      Text            =   "00000000"
      Top             =   1260
      Width           =   1380
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "AbonoRet.frx":0071
      DataSource      =   "AdoFactura"
      Height          =   360
      Left            =   105
      TabIndex        =   11
      Top             =   1260
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "00000000"
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
   Begin MSDataListLib.DataCombo DCRecibo 
      Bindings        =   "AbonoRet.frx":008A
      DataSource      =   "AdoRecibo"
      Height          =   360
      Left            =   7350
      TabIndex        =   7
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "99999999"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "AbonoRet.frx":00A2
      DataSource      =   "AdoCliente"
      Height          =   360
      Left            =   1575
      TabIndex        =   9
      Top             =   525
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   1995
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   105
      Top             =   4620
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Salir"
      Height          =   855
      Left            =   8820
      Picture         =   "AbonoRet.frx":00BB
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1050
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   8820
      Picture         =   "AbonoRet.frx":0985
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   3675
      TabIndex        =   3
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "AbonoRet.frx":0DC7
      DataSource      =   "AdoDetAcomp"
      Height          =   360
      Left            =   5565
      TabIndex        =   5
      Top             =   105
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "FA"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   8400
      Top             =   4620
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
   Begin MSAdodcLib.Adodc AdoRecibo 
      Height          =   330
      Left            =   1995
      Top             =   4935
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
      Caption         =   "Recibo"
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   105
      Top             =   4935
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
      Caption         =   "DetAcomp"
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicio 
      Height          =   330
      Left            =   5985
      Top             =   4935
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "RetIvaServicio"
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   3570
      Top             =   4935
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "RetFuente"
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
   Begin MSAdodcLib.Adodc AdoCodRet 
      Height          =   330
      Left            =   3885
      Top             =   4620
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
      Caption         =   "CodRet"
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
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Autorización"
      Height          =   330
      Left            =   5670
      TabIndex        =   18
      Top             =   945
      Width           =   3060
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Serie"
      Height          =   330
      Left            =   4830
      TabIndex        =   16
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retención No. "
      Height          =   330
      Left            =   3465
      TabIndex        =   14
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label LabelSaldo 
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
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   1575
      TabIndex        =   13
      Top             =   1260
      Width           =   1905
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
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
      Left            =   1575
      TabIndex        =   12
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &TIPO:"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4935
      TabIndex        =   4
      Top             =   105
      Width           =   645
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Grupo No."
      Height          =   330
      Left            =   6300
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F&actura No."
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA:"
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &CLIENTE"
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   525
      Width           =   1485
   End
End
Attribute VB_Name = "AbonoRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EsNotaDeVenta As Boolean
Dim AutorizacionFA As String
Dim SerieFA As String

Public Sub GrabarAbonosRetenciones(CtaCierre As String, _
                                   DetalleCta As String, _
                                   Cheq_Dep As String, _
                                   DetTP As String, _
                                   Factura As Long, _
                                   Valor As Currency, _
                                   FechaAbono As String, _
                                   Establecimiento, _
                                   Emision As String, _
                                   Autorizacion As String, _
                                   Porcentaje As Byte)
  If Valor > 0 Then
     If Cheq_Dep = Ninguno And DiarioCaja > 0 Then Cheq_Dep = CStr(DiarioCaja)
     SetAdoAddNew "Trans_Abonos"
     SetAdoFields "T", Normal '
     SetAdoFields "TP", DetTP '
     SetAdoFields "Fecha", FechaAbono '
     SetAdoFields "Recibo_No", Format$(DiarioCaja, "0000000")
     SetAdoFields "Cta", CtaCierre '
     SetAdoFields "Cta_CxP", Cta_Cobrar
     SetAdoFields "Factura", Factura '
     SetAdoFields "Autorizacion", AutorizacionFA
     SetAdoFields "Serie", SerieFA
     SetAdoFields "CodigoC", CodigoCliente
     SetAdoFields "Abono", Valor '
     SetAdoFields "Banco", UCaseStrg(DetalleCta) '
     SetAdoFields "Cheque", Format$(Val(Cheq_Dep), "0000000")
     SetAdoFields "Codigo_Inv", CodigoP
     SetAdoFields "Comprobante", Autorizacion
     SetAdoFields "EstabRetencion", Establecimiento
     SetAdoFields "PtoEmiRetencion", Emision
     SetAdoFields "Porc", Porcentaje
     SetAdoFields "CodigoU", CodigoUsuario
     SetAdoFields "Item", NumEmpresa
     SetAdoUpdate
  End If
End Sub

Public Sub Carga_ConceptosRetencion(MBFecha As String)
Dim FechaCodAir As String
  If Not IsDate(MBFecha) Then MBFecha = "01/01/2000"
  FechaCodAir = BuscarFecha(MBFecha)
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT * " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCCodRet, AdoCodRet, sSQL, "Codigo"
  DCConceptoRetV = "329"
End Sub

Private Sub ChRetB_Click()
  If ChRetB.value = 1 Then FrmIVAB.Visible = True Else FrmIVAB.Visible = False
End Sub

Private Sub ChRetF_Click()
  If ChRetF.value = 1 Then FrmRet.Visible = True Else FrmRet.Visible = False
End Sub

Private Sub ChRetS_Click()
    If ChRetS.value = 1 Then FrmIVAS.Visible = True Else FrmIVAS.Visible = False
End Sub

Private Sub Command1_Click()
  If CFechaLong(MBFecha) < CFechaLong(FechaCorte) Then
     MsgBox "No se puede grabar abonos con fecha inferior a la emision de la factura"
     MBFecha.SetFocus
  Else
    TextoValido TextBanco
    TextoValido TextCheqNo
    Mensajes = "Esta Seguro que desea grabar estos pagos."
    Titulo = "Formulario de Grabación."
    If BoxMensaje = vbYes Then
       Calculo_Saldo
       TotalAbonos = Total_RetIVAB + Total_RetIVAS + Total_Ret
       SaldoDisp = Saldo - TotalAbonos
       Control_Procesos "P", "Abono de " & TipoFactura & ". No. " & Factura_No & "Por Ret " & Format$(TotalAbonos, "#,##0.00")
       FechaTexto = FechaSistema
       If CheqRecibo.value = 1 Then
          DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
       Else
          DiarioCaja = Val(TxtRecibo)
       End If
       Codigo1 = Format$(Val(TrimStrg(MidStrg(TextCheqNo, 1, 3))), "000")
       Codigo2 = Format$(Val(TrimStrg(MidStrg(TextCheqNo, 4, 3))), "000")
       If ChRetB.value <> 0 Then
          Cta_Ret_IVA = SinEspaciosIzq(DCRetIBienes)
          GrabarAbonosRetenciones Cta_Ret_IVA, "RETENCION IVA BIENES", TextCompRet, TipoFactura, Factura_No, Total_RetIVAB, MBFecha, Codigo1, Codigo2, TextBanco, Val(CBienes)
       End If
       If ChRetS.value <> 0 Then
          Cta_Ret_IVA = SinEspaciosIzq(DCRetISer)
          GrabarAbonosRetenciones Cta_Ret_IVA, "RETENCION IVA SERVICIO", TextCompRet, TipoFactura, Factura_No, Total_RetIVAS, MBFecha, Codigo1, Codigo2, TextBanco, Val(CServicio)
       End If
       If ChRetF.value <> 0 Then
          Cta_Ret = SinEspaciosIzq(DCRetFuente)
          GrabarAbonosRetenciones Cta_Ret, "RETENCION FUENTE - " & DCCodRet, TextCompRet, TipoFactura, Factura_No, Total_Ret, MBFecha, Codigo1, Codigo2, TextBanco, Val(TextPorc)
       End If
       T = "P"
       If SaldoDisp <= 0 Then
          T = "C"
          SaldoDisp = 0
       End If
       sSQL = "UPDATE Facturas " _
            & "SET Saldo_MN = " & SaldoDisp & ",T = '" & T & "' " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Factura = " & Factura_No & " " _
            & "AND TC = '" & TipoFactura & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodigoC = '" & CodigoCliente & "' "
       Ejecutar_SQL_SP sSQL
       TextCajaMN = "0.00"
       TextCajaME = "0.00"
       TextRet = "0.00"
       TextRetIVA = "0.00"
       TextCheqNo = ""
       TextCheque = "0.00"
       TextBaucher = ""
       TextTotalBaucher = "0.00"
       TextRecibido = "0.00"
       TextInteres = "0.00"
       RatonNormal
    End If
  End If
End Sub

Private Sub Command2_Click()
   Unload AbonoRetencion
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
   If KeyCode = vbKeyEscape Then Command2.SetFocus
End Sub

Private Sub DCCliente_LostFocus()
  With AdoCliente.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente Like '" & DCCliente.Text & "' ")
       If Not .EOF Then
          CodigoCliente = .fields("Codigo")
          NombreCliente = DCCliente.Text
          DireccionCli = .fields("Direccion")
          'Grupo_No = .Fields("Grupo")
          Listar_Facturas_P
          If Evaluar Then DCFactura.Text = Factura_No
       Else
          MsgBox "Cliente no Asignado"
       End If
   Else
       MsgBox "No existen datos"
   End If
  End With
End Sub

Private Sub DCFactura_GotFocus()
  Total_RetIVAB = 0
  Total_RetIVAS = 0
  Total_Ret = 0
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then DCCliente.SetFocus
End Sub

Private Sub DCFactura_LostFocus()
  Factura_No = 0: Saldo = 0: Cotizacion = 0: TotalDolar = 0: Saldo_ME = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura Like " & Val(DCFactura) & " ")
       If Not .EOF Then
          FechaCorte = .fields("Fecha")
          Factura_No = .fields("Factura")
          Cta_Cobrar = .fields("Cta_CxP")
          Saldo = Redondear(.fields("Saldo_MN"), 2)
          Saldo_ME = Redondear(.fields("Saldo_ME"), 2)
          TipoFactura = .fields("TC")
          TotalDolar = .fields("Total_ME")
          Cotizacion = .fields("Cotizacion")
          AutorizacionFA = .fields("Autorizacion")
          SerieFA = .fields("Serie")
          Total_Ret = 0
          Total_RetIVAB = 0
          Total_RetIVAS = 0
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
          'TextRet.SetFocus
          AbonoRetencion.Caption = "INGRESO DE CAJA (" & TipoFactura & ")"
       End If
    End If
  End With
  Calculo_Saldo
End Sub

Private Sub DCRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCRecibo_LostFocus()
  Listar_Clientes_F
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  Grupos_Pendientes
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextRet
  ControlEsNumerico TextRetIVAB
  ControlEsNumerico TextRetIVAS
  Grupos_Pendientes
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  If CheqRecibo.value = 1 Then TxtRecibo = Format$(DiarioCaja, "0000000") Else TxtRecibo = ""
  EsNotaDeVenta = False
  Mifecha = BuscarFecha(FechaSistema)
  FechaTexto = FechaSistema
  CBienes.Clear
  CBienes.AddItem "0"
  CBienes.AddItem "30"
  CBienes.AddItem "100"
  CBienes.Text = "0"
  
  CServicio.Clear
  CServicio.AddItem "0"
  CServicio.AddItem "70"
  CServicio.AddItem "100"
  CServicio.Text = "0"
  
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC <> 'OP' " _
       & "AND T = 'P' " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDB_Combo DCTipo, AdoDetAcomp, sSQL, "TC"
  
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CF' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
  
 'Carga los Conceptos de retención IVA Servicios al DataCombo
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CI' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetISer, AdoRetIvaServicio, sSQL, "Cuentas"
 'Carga los Conceptos de retención IVA Bienes al DataCombo
  sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CI' " _
       & "AND DG = 'D' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCRetIBienes, AdoRetIvaBienes, sSQL, "Cuentas"
  
' Abonamos la factura
  If TipoFactura = "NV" Then EsNotaDeVenta = True
  If CodigoCliente = "" Then CodigoCliente = Ninguno
  If Factura_No <= 0 Then Factura_No = 0
  'MsgBox "...."
  If Len(TipoDoc) > 1 Then DCTipo = TipoDoc
  If Len(Grupo_No) > 1 Then DCRecibo = Grupo_No
  Listar_Clientes_F
  Listar_Facturas_P
  AbonoRetencion.Caption = "INGRESO DE CAJA POR RETENCIONES (" & TipoFactura & ")"
  If FechaTexto = "" Or FechaTexto = Ninguno Then FechaTexto = FechaSistema
  MBFecha = FechaTexto
  RatonNormal
  CheqRecibo.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm AbonoRetencion
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoRecibo
   ConectarAdodc AdoCliente
   ConectarAdodc AdoDetAcomp
   ConectarAdodc AdoCodRet
   ConectarAdodc AdoRetFuente
   ConectarAdodc AdoRetIvaBienes
   ConectarAdodc AdoRetIvaServicio
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha, True
  Carga_ConceptosRetencion MBFecha
End Sub

Private Sub TextBanco_GotFocus()
  MarcarTexto TextBanco
End Sub

Private Sub TextBanco_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextBanco_LostFocus()
  TextoValido TextBanco
  If Len(TextBanco) < 10 Then TextBanco = Format$(Val(TextBanco), "0000000000")
End Sub

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
  If Val(TextCheqNo) < 1001 Then TextCheqNo = "001001"
  TextCheqNo = Format$(Val(TextCheqNo), "000000")
End Sub

Private Sub TextCompRet_GotFocus()
  MarcarTexto TextCompRet
End Sub

Private Sub TextCompRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextCompRet_LostFocus()
  TextCompRet = Format$(Val(TextCompRet), "000000000")
End Sub

Private Sub TextPorc_GotFocus()
  MarcarTexto TextPorc
End Sub

Private Sub TextPorc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRet_GotFocus()
  MarcarTexto TextRet
End Sub

Private Sub TextRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextRet_LostFocus()
  TextoValido TextRet, True
  Total_Ret = Redondear(Val(CCur(TextRet)), 2)
  TextRet = Format$(Total_Ret, "#,##0.00")
  Calculo_Saldo
End Sub

Private Sub TextRetIVAB_GotFocus()
  MarcarTexto TextRetIVAB
End Sub

Private Sub TextRetIVAB_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextRetIVAB_LostFocus()
  TextoValido TextRetIVAB, True
  Total_RetIVAB = Redondear(Val(CCur(TextRetIVAB)), 2)
  TextRetIVAB = Format$(Total_RetIVAB, "#,##0.00")
  Calculo_Saldo
End Sub

Private Sub TextRetIVAS_GotFocus()
  MarcarTexto TextRetIVAS
End Sub

Private Sub TextRetIVAS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextRetIVAS_LostFocus()
  TextoValido TextRetIVAS, True
  Total_RetIVAS = Redondear(Val(CCur(TextRetIVAS)), 2)
  TextRetIVAS = Format$(Total_RetIVAS, "#,##0.00")
  Calculo_Saldo
End Sub

Public Sub Grupos_Pendientes()
  TipoFactura = DCTipo.Text
  If TipoFactura = "" Then TipoFactura = Ninguno
  sSQL = "SELECT C.Grupo " _
       & "FROM Clientes As C,Facturas As F " _
       & "WHERE F.T = 'P' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND F.CodigoC = C.Codigo " _
       & "GROUP BY C.Grupo " _
       & "ORDER BY C.Grupo "
  SelectDB_Combo DCRecibo, AdoRecibo, sSQL, "Grupo"
  'Label19.Caption = " TIPO PROCESO (" & TipoFactura & ")"
  If AdoRecibo.Recordset.RecordCount <= 0 Then CheqRecibo.SetFocus
End Sub

Public Sub Listar_Facturas_P()
  TipoFactura = DCTipo
  SQL1 = "SELECT * " _
       & "FROM Facturas " _
       & "WHERE CodigoC = '" & CodigoCliente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND T = '" & Pendiente & "' " _
       & "AND TC = '" & TipoFactura & "' "
  If Factura_No > 0 Then SQL1 = SQL1 & "AND Factura = " & Factura_No & " "
  SQL1 = SQL1 & "ORDER BY Factura "
  SelectDB_Combo DCFactura, AdoFactura, SQL1, "Factura"
  'MsgBox SQL1
  If AdoFactura.Recordset.RecordCount <= 0 Then
     'MsgBox "No existen Valores pendientes por Cobrar"
     Factura_No = 0
     CheqRecibo.SetFocus
  End If
End Sub

Public Sub Listar_Clientes_F()
  Codigo1 = DCRecibo
  If Codigo1 = "" Then Codigo1 = Ninguno
  sSQL = "SELECT Grupo,Cliente,Codigo,CI_RUC,C.Direccion,C.Telefono " _
       & "FROM Clientes As C,Facturas As F " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T = 'P' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND C.Codigo = F.CodigoC "
  If Codigo1 <> Ninguno Then sSQL = sSQL & "AND C.Grupo = '" & Codigo1 & "' "
  sSQL = sSQL & "GROUP BY Grupo,Cliente,Codigo,CI_RUC,C.Direccion,C.Telefono " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
  If AdoCliente.Recordset.RecordCount <= 0 Then CheqRecibo.SetFocus
End Sub

Private Sub TxtRecibo_GotFocus()
  MarcarTexto TxtRecibo
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format$(Val(TxtRecibo), "0000000")
End Sub

Public Sub Calculo_Saldo()
   SaldoDisp = Saldo - Total_Ret - Total_RetIVAB - Total_RetIVAS
End Sub
