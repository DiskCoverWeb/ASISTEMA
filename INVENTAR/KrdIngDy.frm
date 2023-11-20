VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Kard_Ing_DYE 
   Caption         =   "HOJA DE COSTEOS POR COMPRA"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   Icon            =   "KrdIngDy.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11070
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Code39Clt1 
      Height          =   480
      Left            =   2835
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   65
      Top             =   6615
      Width           =   1200
   End
   Begin VB.Frame FrmCambios 
      BackColor       =   &H00C0FFFF&
      Caption         =   " CAMBIO DE VALORES "
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
      Height          =   1170
      Left            =   105
      TabIndex        =   48
      Top             =   3465
      Visible         =   0   'False
      Width           =   10830
      Begin VB.TextBox TxtCant1 
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
         IMEMode         =   3  'DISABLE
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   50
         Text            =   "KrdIngDy.frx":0442
         Top             =   630
         Width           =   1170
      End
      Begin VB.TextBox TxtFOB1 
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
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "KrdIngDy.frx":0444
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox TxtCom1 
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
         Left            =   2520
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   54
         Text            =   "KrdIngDy.frx":0448
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox TxtTransUnit1 
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
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   56
         Text            =   "KrdIngDy.frx":044A
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox TxtCIF1 
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
         IMEMode         =   3  'DISABLE
         Left            =   6510
         MaxLength       =   10
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   60
         Text            =   "KrdIngDy.frx":044C
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox TxtUtil1 
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
         Left            =   7770
         MultiLine       =   -1  'True
         TabIndex        =   62
         Text            =   "KrdIngDy.frx":044E
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox TxtPVP1 
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
         Left            =   9030
         MultiLine       =   -1  'True
         TabIndex        =   64
         Text            =   "KrdIngDy.frx":0450
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CANTIDAD"
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
         TabIndex        =   49
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRECIO FOB"
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
         Left            =   1260
         TabIndex        =   51
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " COM. %"
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
         Left            =   2520
         TabIndex        =   53
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TRANS. UNIT."
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
         Left            =   3360
         TabIndex        =   55
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PRECIO CIF"
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
         Left            =   6510
         TabIndex        =   59
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TRANSP. TOTAL"
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
         Left            =   4830
         TabIndex        =   57
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " UTILIDAD %"
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
         Left            =   7770
         TabIndex        =   61
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.V.P."
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
         Left            =   9030
         TabIndex        =   63
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   4830
         TabIndex        =   58
         Top             =   630
         Width           =   1695
      End
   End
   Begin VB.TextBox TextConcepto 
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
      MaxLength       =   100
      TabIndex        =   12
      Top             =   735
      Width           =   9465
   End
   Begin VB.TextBox TxtTranspor 
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
      Left            =   9135
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "KrdIngDy.frx":0452
      Top             =   105
      Width           =   1800
   End
   Begin VB.TextBox TextTotal 
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
      Left            =   9135
      MultiLine       =   -1  'True
      TabIndex        =   38
      Text            =   "KrdIngDy.frx":0454
      Top             =   2595
      Width           =   1800
   End
   Begin VB.TextBox TxtUtil 
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
      Left            =   7875
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "KrdIngDy.frx":0456
      Top             =   2625
      Width           =   1275
   End
   Begin VB.TextBox TxtCIF 
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
      IMEMode         =   3  'DISABLE
      Left            =   6615
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "KrdIngDy.frx":0458
      Top             =   2625
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   5355
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
   Begin VB.TextBox TxtSubTotal 
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
      Left            =   5775
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "KrdIngDy.frx":045A
      Top             =   6825
      Width           =   1800
   End
   Begin MSDataListLib.DataCombo DCTInv 
      Bindings        =   "KrdIngDy.frx":045C
      DataSource      =   "AdoTInv"
      Height          =   315
      Left            =   105
      TabIndex        =   13
      Top             =   1155
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin VB.OptionButton OpcIVA 
      BackColor       =   &H00FF8080&
      Caption         =   "Con IVA"
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
      Left            =   8505
      TabIndex        =   17
      Top             =   1470
      Value           =   -1  'True
      Width           =   1170
   End
   Begin VB.OptionButton OpcX 
      BackColor       =   &H00FF8080&
      Caption         =   "Sin IVA"
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
      Left            =   9660
      TabIndex        =   18
      Top             =   1470
      Width           =   1275
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "KrdIngDy.frx":0472
      DataSource      =   "AdoBodega"
      Height          =   315
      Left            =   5040
      TabIndex        =   16
      Top             =   1470
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16761024
      Text            =   "DC"
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
   Begin MSDataGridLib.DataGrid DGKardex 
      Bindings        =   "KrdIngDy.frx":048A
      Height          =   3375
      Left            =   105
      TabIndex        =   39
      Top             =   3045
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   16761024
      BorderStyle     =   0
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "KrdIngDy.frx":04A2
      DataSource      =   "AdoInv"
      Height          =   765
      Left            =   105
      TabIndex        =   14
      Top             =   1410
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   1349
      _Version        =   393216
      Style           =   1
      BackColor       =   16744576
      ForeColor       =   16777215
      Text            =   "Productos"
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
   Begin VB.TextBox TextDesc1 
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
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "KrdIngDy.frx":04B7
      Top             =   2625
      Width           =   1485
   End
   Begin VB.TextBox TextDesc 
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
      Left            =   2520
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "KrdIngDy.frx":04B9
      Top             =   2625
      Width           =   855
   End
   Begin VB.TextBox TextVUnit 
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
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "KrdIngDy.frx":04BB
      Top             =   2625
      Width           =   1275
   End
   Begin VB.TextBox TextEntrada 
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
      IMEMode         =   3  'DISABLE
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "KrdIngDy.frx":04BF
      Top             =   2625
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo DCBenef 
      Bindings        =   "KrdIngDy.frx":04C1
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   1470
      TabIndex        =   7
      Top             =   420
      Width           =   4950
      _ExtentX        =   8731
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
   Begin MSAdodcLib.Adodc AdoCtaObra 
      Height          =   330
      Left            =   2310
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
      Caption         =   "CtaObra"
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
      Left            =   2310
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
   Begin MSAdodcLib.Adodc AdoKardex 
      Height          =   330
      Left            =   4320
      Top             =   4365
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
      Caption         =   "Kardex"
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   4305
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
      Caption         =   "Inv"
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   315
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   4320
      Top             =   4050
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
   Begin VB.TextBox TextOrden 
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
      IMEMode         =   3  'DISABLE
      Left            =   7560
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "KrdIngDy.frx":04D8
      Top             =   420
      Width           =   1590
   End
   Begin VB.TextBox TextIVA 
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
      Left            =   7665
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "KrdIngDy.frx":04DA
      Top             =   6825
      Width           =   1590
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
      Height          =   750
      Left            =   1470
      Picture         =   "KrdIngDy.frx":04DC
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6510
      Width           =   1170
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
      Height          =   750
      Left            =   210
      Picture         =   "KrdIngDy.frx":0DA6
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6510
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoTInv 
      Height          =   330
      Left            =   315
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
      Caption         =   "TInv"
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   315
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
      Caption         =   "Bodega"
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
   Begin MSAdodcLib.Adodc AdoRet 
      Height          =   330
      Left            =   2310
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
      Caption         =   "Ret"
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
   Begin MSAdodcLib.Adodc AdoIVA 
      Height          =   330
      Left            =   315
      Top             =   4725
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
      Caption         =   "IVA"
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
   Begin MSDataListLib.DataCombo DCDiario 
      Bindings        =   "KrdIngDy.frx":11E8
      DataSource      =   "AdoRet"
      Height          =   315
      Left            =   2835
      TabIndex        =   1
      Top             =   105
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Diario"
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
      Left            =   2310
      Top             =   4725
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONCEPTO"
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
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comprobante de Diario No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2745
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TRANSPORTE TOTAL"
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
      Left            =   6615
      TabIndex        =   4
      Top             =   105
      Width           =   2535
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   4830
      TabIndex        =   32
      Top             =   2625
      Width           =   1800
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRECIO VENTA"
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
      Left            =   9135
      TabIndex        =   37
      Top             =   2310
      Width           =   1800
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000005&
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
      Left            =   6195
      TabIndex        =   20
      Top             =   1890
      Width           =   2220
   End
   Begin VB.Label LabelUnidad 
      BackColor       =   &H80000005&
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
      Left            =   9450
      TabIndex        =   22
      Top             =   1890
      Width           =   1485
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GUIA No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   6405
      TabIndex        =   8
      Top             =   420
      Width           =   1170
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UTILIDAD %"
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
      Left            =   7875
      TabIndex        =   35
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TRANSP. TOTAL"
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
      Left            =   4830
      TabIndex        =   31
      Top             =   2310
      Width           =   1800
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL ASIENTO"
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
      Left            =   5775
      TabIndex        =   47
      Top             =   6510
      Width           =   1800
   End
   Begin VB.Label LabelProducto 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   1155
      Width           =   5895
   End
   Begin VB.Label Label2 
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
      Left            =   9135
      TabIndex        =   10
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRECIO CIF"
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
      Left            =   6615
      TabIndex        =   33
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD"
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
      Left            =   8505
      TabIndex        =   21
      Top             =   1890
      Width           =   960
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      TabIndex        =   43
      Top             =   6825
      Width           =   1590
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TRANS. UNIT."
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
      Left            =   3360
      TabIndex        =   29
      Top             =   2310
      Width           =   1485
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COM. %"
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
      Left            =   2520
      TabIndex        =   27
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRECIO FOB"
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
      Left            =   1260
      TabIndex        =   25
      Top             =   2310
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CANTIDAD"
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
      TabIndex        =   23
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIFERENCIA"
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
      TabIndex        =   42
      Top             =   6510
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
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
      Left            =   7665
      TabIndex        =   40
      Top             =   6510
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
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
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PROVEEDOR"
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
      Top             =   420
      Width           =   1380
   End
End
Attribute VB_Name = "Kard_Ing_DYE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLDYE  As String

Private Sub Command1_Click()
  RatonReloj
  Total_RetCta = 0
  RatonNormal
  Mensajes = "Seguro de Grabar?"
  Titulo = "GRABACION DEL COMPROBANTE"
  If BoxMensaje = vbYes Then
  RatonReloj
  FechaTexto = MBFechaI.Text
  sSQL = "DELETE * " _
       & "FROM Asiento " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "DELETE * " _
       & "FROM Asiento_R " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  
  sSQL = "SELECT * " _
       & "FROM Asiento_K " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND T_No = " & Trans_No & " " _
       & "ORDER BY CTA_INVENTARIO,CONTRA_CTA "
  Select_Adodc_Grid DGKardex, AdoKardex, sSQL, 4
  RatonReloj
  Si_No = True
  CodigoBenef = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Cliente Like '" & Cadena & "' ")
       If Not .EOF Then CodigoBenef = .Fields("Codigo")
   End If
  End With
  Valor = 0
  TotalInventario
  FechaTexto = MBFechaI.Text
  TextoValido TextOrden, , True
  'MsgBox AdoKardex.Recordset.RecordCount
  'MsgBox CodigoBenef
  If CodigoBenef <> Ninguno Then
     SetAdoAddNew "Asiento_SC"
     SetAdoFields "Factura", Val(TextOrden)
     SetAdoFields "DH", "2"
     SetAdoFields "TM", "1"
     SetAdoFields "Codigo", CodigoBenef
     SetAdoFields "Valor", Total_ME
     SetAdoFields "FECHA_V", FechaTexto
     SetAdoFields "Cta", Contra_Cta
     SetAdoFields "TC", "P"
     SetAdoFields "T_No", Trans_No
     SetAdoFields "SC_No", Asiento
     SetAdoUpdate
     Asiento = Asiento + 1
     sSQL = "SELECT * " _
          & "FROM Asiento " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Select_Adodc AdoAsientos, sSQL
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          CodigoInv = .Fields("Codigo_Inv")
          Cta_Inventario = .Fields("CTA_INVENTARIO")
          Contra_Cta = .Fields("CONTRA_CTA")
          Total = 0: ValorDH = 0
          Asiento = 1
         'Llenamos los datos ingresados al Kardex
          Do While Not .EOF
             If Cta_Inventario <> .Fields("CTA_INVENTARIO") Then
                InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
                CodigoInv = .Fields("Codigo_Inv")
                Cta_Inventario = .Fields("CTA_INVENTARIO")
                Contra_Cta = .Fields("CONTRA_CTA")
                ValorDH = 0
             End If
             Total = Total + .Fields("VALOR_TOTAL")
             ValorDH = ValorDH + .Fields("VALOR_TOTAL")
            .MoveNext
          Loop
          InsertarAsientos AdoAsientos, Cta_Inventario, 0, ValorDH, 0
          InsertarAsientos AdoAsientos, Contra_Cta, 0, 0, ValorDH
          Debe = 0: Haber = 0
          Do While Not .EOF
             Debe = Debe + .Fields("DEBE")
             Haber = Haber + .Fields("HABER")
            .MoveNext
          Loop
          If (Debe - Haber) <> 0 Then MsgBox "Verifique el comprobante, no cuadra por: " & Round(Debe - Haber, 2)
         .MoveFirst
          RatonReloj
          Co.T = Normal
          Co.TP = CompDiario
          Co.Fecha = FechaTexto
          Co.Numero = NumComp
          Co.Concepto = TextConcepto
          Co.CodigoB = CodigoBenef
          Co.Efectivo = CCur(TxtTranspor)
          Co.Monto_Total = Total - Total_RetCta
          Co.Usuario = CodigoUsuario
          Co.T_No = Trans_No
          Co.Item = NumEmpresa
          If TextOrden <> Ninguno Then Co.Concepto = Co.Concepto & ", Orden No. " & TextOrden
          'MsgBox Total_ME
          GrabarComprobante Co
          ImprimirComprobantesDe False, Co
          Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Mensajes = "Imprimir Copia de Nota de Entrada/Salida"
          Titulo = "COPIA DE NOTA"
          If BoxMensaje = vbYes Then Imprimir_Nota_Inventario AdoBenef, AdoKardex, NumComp, TextOrden.Text, "CD", FechaTexto, FechaTexto, Total
          Unload Kard_Ing_DYE
      End If
     End With
  Else
     MsgBox "Beneficiario no asignado"
  End If
  End If
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload Kard_Ing_DYE
End Sub

Private Sub DCBenef_LostFocus()
  InvImp = False
  Label2.Caption = " LOCAL"
  CodigoBenef = Ninguno
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
       Cadena = DCBenef.Text
      .MoveFirst
      .Find ("Cliente Like '" & Cadena & "' ")
       If Not .EOF Then
          CodigoBenef = .Fields("Codigo")
          CodigoCliente = .Fields("Codigo")
          TipoDoc = .Fields("TD")
          InvImp = .Fields("Importaciones")
          If InvImp Then Label2.Caption = " IMPORTACION"
          Si_No = True
          If TipoDoc = "R" Then Si_No = False
       Else
          Si_No = False
       End If
   End If
  End With
  TextConcepto.Text = DCBenef.Text & " "
End Sub

Private Sub DCBodega_GotFocus()
  LabelCodigo.Caption = CodigoInv
  LabelUnidad.Caption = Unidad
  LabelProducto.Caption = Producto
End Sub

Private Sub DCBodega_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBodega_LostFocus()
  SaldoAnterior = 0: ValorUnitAnt = 0: ValorUnit = 0: Cantidad = 0
  Stock_Actual_Inventario MBFechaI, CodigoInv
  Precio = ValorUnit
  TextVUnit.Text = Format(Precio, "#,##0.000000")
End Sub

Private Sub DCDiario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCDiario_LostFocus()
  Contador = 0
  If Val(DCDiario.Text) > 0 Then
     sSQL = "DELETE * " _
          & "FROM Asiento_K " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "AND T_No = " & Trans_No & " "
     Ejecutar_SQL_SP sSQL
     sSQL = SQLDYE
    'SQLDec = "VALOR_UNIT 4|,VALOR_TOTAL 4|,CANTIDAD 4|,SALDO 4|."
     Select_Adodc_Grid DGKardex, AdoKardex, sSQL
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TP = 'CD' " _
          & "AND Numero = " & Val(DCDiario.Text) & " " _
          & "ORDER BY Codigo_Inv,K.ID "
     Select_Adodc AdoAux, sSQL
     Contador = 0
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          CodigoBenef = .Fields("Codigo_P")
          TextOrden = .Fields("Orden_No")
          Cod_Bodega = .Fields("CodBodega")
          Mifecha = .Fields("Fecha")
          Numero = .Fields("Numero")
          Sumatoria = 0
          Do While Not .EOF
             ValorUnit = (.Fields("Precio_FOB") * .Fields("Comision")) + .Fields("Trans_Unit") + .Fields("Precio_FOB")
             Entrada = .Fields("Entrada")
             SubTotal = .Fields("Trans_Unit") * Entrada
             Sumatoria = Sumatoria + SubTotal
           ' Llenamos el ultimo saldo del kardex
             If Entrada > 0 And ValorUnit > 0 Then
                ValorTotal = ValorUnit * Entrada
                SetAddNew AdoKardex
                SetFields AdoKardex, "DH", "1"
                SetFields AdoKardex, "CANT_ES", Entrada
                SetFields AdoKardex, "CODIGO_INV", .Fields("Codigo_Inv")
                SetFields AdoKardex, "PRODUCTO", .Fields("Producto")
                SetFields AdoKardex, "VALOR_FOB", .Fields("Precio_FOB")
                SetFields AdoKardex, "TRANS_UNIT", .Fields("Trans_Unit")
                SetFields AdoKardex, "COMIS", .Fields("Comision")
                SetFields AdoKardex, "UTIL", .Fields("Utilidad")
                SetFields AdoKardex, "VALOR_UNIT", .Fields("Valor_Unitario")
                SetFields AdoKardex, "VALOR_TOTAL", .Fields("Valor_Total")
                SetFields AdoKardex, "CTA_INVENTARIO", .Fields("Cta_Inv")
                SetFields AdoKardex, "PVP", .Fields("PVP")
                SetFields AdoKardex, "PRECIO_CIF", ValorUnit
                SetFields AdoKardex, "TRANS_TOTAL", SubTotal
                SetFields AdoKardex, "CONTRA_CTA", .Fields("Contra_Cta")
                SetFields AdoKardex, "ORDEN", Val(CCur(TextOrden))
                SetFields AdoKardex, "Codigo_B", CodigoBenef
                SetFields AdoKardex, "CodBod", Cod_Bodega
                SetFields AdoKardex, "Item", NumEmpresa
                SetFields AdoKardex, "CodigoU", CodigoUsuario
                SetFields AdoKardex, "T_No", Trans_No
                SetFields AdoKardex, "A_No", Contador
                SetUpdate AdoKardex
                Contador = Contador + 1
             End If
            .MoveNext
          Loop
          TotalInventario
          If AdoBenef.Recordset.RecordCount > 0 Then
             AdoBenef.Recordset.MoveFirst
             AdoBenef.Recordset.Find ("Codigo Like '" & CodigoBenef & "' ")
             If Not AdoBenef.Recordset.EOF Then
                DCBenef = AdoBenef.Recordset.Fields("Cliente")
                CodigoCliente = AdoBenef.Recordset.Fields("Codigo")
                TipoDoc = AdoBenef.Recordset.Fields("TD")
                InvImp = AdoBenef.Recordset.Fields("Importaciones")
                If InvImp Then Label2.Caption = " IMPORTACION"
                Si_No = True
                If TipoDoc = "R" Then Si_No = False
             End If
          End If
          MBFechaI = Mifecha
          TxtTranspor = Sumatoria
          MBFechaI.SetFocus
      Else
          MsgBox "No existe datos"
      End If
     End With
  Else
     MsgBox "Usted va ha ingresar una nueva compra"
  End If
  MBFechaI.SetFocus
End Sub

Private Sub DCInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Command1.SetFocus
End Sub

Private Sub DCInv_LostFocus()
  CodigoInv = DCInv.Text
  Si_No = False
  'MsgBox DCInv.Text
  With AdoInv.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Producto Like '" & CodigoInv & "' ")
       If Not .EOF Then
          Si_No = .Fields("IVA")
          Unidad = .Fields("Unidad")
          CodigoInv = .Fields("Codigo_Inv")
          Producto = .Fields("Producto")
          Cta_Inventario = .Fields("Cta_Inventario")
          Contra_Cta = .Fields("Cta_Proveedor")
          Contra_Cta1 = .Fields("Cta_Costo_Venta")
          DCBodega.SetFocus
       Else
         .MoveFirst
         .Find ("Codigo_Barra Like '" & CodigoInv & "' ")
          If Not .EOF Then
             Si_No = .Fields("IVA")
             Unidad = .Fields("Unidad")
             CodigoInv = .Fields("Codigo_Inv")
             Producto = .Fields("Producto")
             Cta_Inventario = .Fields("Cta_Inventario")
             Contra_Cta = .Fields("Cta_Proveedor")
             Contra_Cta1 = .Fields("Cta_Costo_Venta")
             DCBodega.SetFocus
          Else
            .MoveFirst
            .Find ("Codigo_Inv Like '" & CodigoInv & "' ")
             If Not .EOF Then
                Si_No = .Fields("IVA")
                Unidad = .Fields("Unidad")
                CodigoInv = .Fields("Codigo_Inv")
                Producto = .Fields("Producto")
                Cta_Inventario = .Fields("Cta_Inventario")
                Contra_Cta = .Fields("Cta_Proveedor")
                Contra_Cta1 = .Fields("Cta_Costo_Venta")
                DCBodega.SetFocus
             Else
                MsgBox "No existe Productos asignados"
                DCInv.SetFocus
             End If
          End If
       End If
    End If
  End With
  If Si_No Then OpcIVA.value = True Else OpcX.value = True
  If Empleados Then OpcX.value = True
  If InvImp Then OpcX.value = True
End Sub

Private Sub DCTInv_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTInv_LostFocus()
  ListarProductos
End Sub

Private Sub DGKardex_BeforeDelete(Cancel As Integer)
  Cancel = DeleteSiNo(AdoKardex)
End Sub

Private Sub DGKardex_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF10 Then
     With AdoKardex.Recordset
      If .RecordCount > 0 Then
          TxtCant1 = .Fields("CANT_ES")
          TxtFOB1 = .Fields("VALOR_FOB")
          TxtCom1 = .Fields("COMIS") * 100
          TxtTransUnit1 = .Fields("TRANS_UNIT")
          Label23.Caption = .Fields("TRANS_TOTAL")
          TxtCIF1 = .Fields("PRECIO_CIF")
          TxtUtil1 = .Fields("UTIL") * 100
          TxtTxtPVP1 = .Fields("PVP")
          FrmCambios.Visible = True
          TxtCant1.SetFocus
      End If
     End With
  End If
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub TextConcepto_GotFocus()
 MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Private Sub TextDesc_GotFocus()
  MarcarTexto TextDesc
End Sub

Private Sub TextDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc_LostFocus()
  TextoValido TextDesc
End Sub

Private Sub TextDesc1_GotFocus()
  MarcarTexto TextDesc1
End Sub

Private Sub TextDesc1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDesc1_LostFocus()
   Label19.Caption = Format(CCur(TextEntrada) * CCur(TextDesc1), "#,##0.00")
End Sub

Private Sub TextEntrada_GotFocus()
  OpcDH = 1
  TotalInventario
  TextVUnit.Text = Format(ValorUnit, "#,##0.0000")
  MarcarTexto TextEntrada
End Sub

Private Sub TextEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextEntrada_LostFocus()
  TextoValido TextEntrada, True, , 0
  Entrada = CCur(TextEntrada.Text)
End Sub

Private Sub Form_Activate()
 SQLDYE = "SELECT CODIGO_INV,PRODUCTO,CANT_ES,VALOR_FOB,COMIS,TRANS_UNIT,TRANS_TOTAL,PRECIO_CIF," _
        & "UTIL,PVP,IVA,CTA_INVENTARIO,CONTRA_CTA,SUBCTA,CodBod,T_No,Item,CodigoU,A_No,TC," _
        & "DH,VALOR_UNIT,VALOR_TOTAL,CANTIDAD,SALDO,P_DESC,P_DESC1,UNIDAD,COD_BAR,Cod_Tarifa," _
        & "Fecha_DUI,No_Refrendo,DUI,ValorEM,Especifico,Consumos,Antidumping,Modernizacion," _
        & "Control,Almacenaje,FODIN,Salvaguarda,Interes,CODIGO_INV1,CodBod1,Codigo_B,ORDEN " _
        & "FROM Asiento_K " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " "
'Consultamos las cuentas de la tabla
 Trans_No = 97
 sSQL = "DELETE * " _
      & "FROM Asiento_K " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND CodigoU = '" & CodigoUsuario & "' " _
      & "AND T_No = " & Trans_No & " "
 Ejecutar_SQL_SP sSQL
 IniciarAsientosAdo AdoAsientos
 TipoDoc = CompDiario
 sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TC = 'I' " _
      & "ORDER BY Codigo_Inv "
 SelectDB_Combo DCTInv, AdoTInv, sSQL, "NomProd"
 ListarProductos
 sSQL = SQLDYE
 'SQLDec = "VALOR_UNIT 4|,VALOR_TOTAL 4|,CANTIDAD 4|,SALDO 4|."
 Select_Adodc_Grid DGKardex, AdoKardex, sSQL
  
 sSQL = "SELECT * " _
      & "FROM Catalogo_Bodegas " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "ORDER BY CodBod "
 SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
 
 sSQL = "SELECT Numero " _
      & "FROM Trans_Kardex " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND TP = 'CD' " _
      & "AND Entrada > 0 " _
      & "AND Precio_FOB > 0 " _
      & "GROUP BY Numero " _
      & "ORDER BY Numero DESC "
 SelectDB_Combo DCDiario, AdoRet, sSQL, "Numero"
 If AdoRet.Recordset.RecordCount <= 0 Then DCDiario.Text = "0"
 FechaValida MBFechaI
 RatonNormal
 Total_IVA = 0
 Label11.Visible = True
 TextIVA.Visible = True
 Label3.Caption = " PROVEEDOR:"
 ListarProveedorUsuario True
 TextOrden.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoInv
  ConectarAdodc AdoRet
  ConectarAdodc AdoIVA
  ConectarAdodc AdoAux
  ConectarAdodc AdoArt
  ConectarAdodc AdoTInv
  ConectarAdodc AdoBenef
  ConectarAdodc AdoKardex
  ConectarAdodc AdoBodega
  ConectarAdodc AdoCtaObra
  ConectarAdodc AdoAsientos
End Sub

Private Sub TextIVA_GotFocus()
  TextIVA.Text = ""
  TotalInventario
End Sub

Private Sub TextIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextIVA_LostFocus()
  TextoValido TextIVA, True
  TotalInventario
End Sub

Private Sub TextOrden_GotFocus()
  MarcarTexto TextOrden
  Cod_Bodega = Ninguno
End Sub

Private Sub TextOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextOrden_LostFocus()
  TextoValido TextOrden, , True
End Sub

Private Sub TextTotal_GotFocus()
  TextTotal.Text = Format(Val(CCur(TextTotal)), "#,##0.0000")
End Sub

Private Sub TextTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextTotal_LostFocus()
   Entrada = CCur(TextEntrada)
   ValorUnit = Val(CCur(TxtCIF))
   Codigo = Leer_Cta_Catalogo(Contra_Cta)
 ' Llenamos el ultimo saldo del kardex
   If Entrada > 0 And ValorUnit > 0 Then
      ValorTotal = ValorUnit * Entrada
      SetAddNew AdoKardex
      SetFields AdoKardex, "DH", "1"
      SetFields AdoKardex, "CODIGO_INV", CodigoInv
      SetFields AdoKardex, "PRODUCTO", Producto
      SetFields AdoKardex, "CANT_ES", Entrada
      SetFields AdoKardex, "VALOR_FOB", Val(CCur(TextVUnit))
      SetFields AdoKardex, "TRANS_UNIT", Val(CCur(TextDesc1))
      SetFields AdoKardex, "TRANS_TOTAL", Val(CCur(Label19.Caption))
      SetFields AdoKardex, "PRECIO_CIF", Val(CCur(TxtCIF))
      SetFields AdoKardex, "COMIS", Val(CCur(TextDesc)) / 100
      SetFields AdoKardex, "UTIL", Val(CCur(TxtUtil)) / 100
      SetFields AdoKardex, "PVP", Val(CCur(TextTotal))
      SetFields AdoKardex, "SUBCTA", Ninguno
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "VALOR_TOTAL", ValorTotal
      SetFields AdoKardex, "CTA_INVENTARIO", Cta_Inventario
      SetFields AdoKardex, "CONTRA_CTA", Contra_Cta
      SetFields AdoKardex, "ORDEN", Val(CCur(TextOrden))
      SetFields AdoKardex, "Codigo_B", CodigoBenef
      SetFields AdoKardex, "CodBod", Cod_Bodega
      SetFields AdoKardex, "Item", NumEmpresa
      SetFields AdoKardex, "CodigoU", CodigoUsuario
      SetFields AdoKardex, "T_No", Trans_No
      SetFields AdoKardex, "A_No", Contador
      SetUpdate AdoKardex
      Contador = Contador + 1
      TotalInventario
      DCInv.SetFocus
   Else
      TotalInventario
   End If
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
  TextoValido TextVUnit, True, , 4
  ValorUnit = CCur(TextVUnit)
  'ValorTotal = ValorUnit * Entrada
  'TextTotal.Text = Format(ValorTotal, "#,##0.0000")
  'TextVUnit.Text = Format(ValorUnit, "#,##0.0000")
End Sub

Public Sub TotalInventario()
Dim TotalInvs As Currency
  Total = 0: Total_IVA = 0: Total_ME = 0
  Saldo = Val(CCur(TxtTranspor))
  With AdoKardex.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .MoveFirst
        Do While Not .EOF
           Total = Total + .Fields("TRANS_TOTAL")
           Total_ME = Total_ME + .Fields("VALOR_TOTAL")
           Total_IVA = Total_IVA + .Fields("IVA")
          .MoveNext
        Loop
       .MoveFirst
    End If
  End With
  'MsgBox Total
  TxtSubTotal.Text = Format(Total_ME, "#,##0.00")
  TextIVA.Text = Format(Total_IVA, "#,##0.00")
  Label1.Caption = Format(Saldo - Total, "#,##0.00")
  If (Saldo - Total) < 0 Then MsgBox "Total Execido, vuelva a recalcular"
End Sub

Public Sub ListarProductos()
 CodigoInv = SinEspaciosIzq(DCTInv)
 sSQL = "SELECT * " _
      & "FROM Catalogo_Productos " _
      & "WHERE Item = '" & NumEmpresa & "' " _
      & "AND Periodo = '" & Periodo_Contable & "' " _
      & "AND MidStrg(Codigo_Inv,1," & CStr(Len(CodigoInv)) & ") = '" & CodigoInv & "' " _
      & "AND LEN(Cta_Inventario) > 1 " _
      & "AND TC = 'P' " _
      & "ORDER BY Producto "
 SelectDB_Combo DCInv, AdoInv, sSQL, "Producto"
End Sub

Public Sub ListarProveedorUsuario(Proveedor As Boolean)
  OpcIVA.Visible = True
  OpcX.Visible = True
  sSQL = "SELECT C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones " _
       & "FROM Clientes As C,Catalogo_CxCxP As CP " _
       & "WHERE CP.TC = 'P' " _
       & "AND CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = CP.Codigo " _
       & "GROUP BY C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,C.TD,CP.Cta,CP.Importaciones " _
       & "ORDER BY C.Cliente "
  SelectDB_Combo DCBenef, AdoBenef, sSQL, "Cliente"
End Sub

Private Sub TxtCant1_GotFocus()
  OpcDH = 1
  MarcarTexto TxtCant1
End Sub

Private Sub TxtCant1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then FrmCambios.Visible = False
End Sub

Private Sub TxtCIF_GotFocus()
  TxtCIF = (Val(CCur(TextVUnit)) * (Val(CCur(TextDesc)) / 100)) + Val(CCur(TextDesc1)) + Val(CCur(TextVUnit))
  ValorTotal = ValorUnit * Entrada
End Sub

Private Sub TxtCIF1_GotFocus()
  TxtCIF1 = (Val(CCur(TxtFOB1)) * (Val(CCur(TxtCom1)) / 100)) + Val(CCur(TxtTransUnit1)) + Val(CCur(TxtFOB1))
  ValorTotal = ValorUnit * Entrada
  MarcarTexto TxtCIF1
End Sub

Private Sub TxtCom1_GotFocus()
   MarcarTexto TxtCom1
End Sub

Private Sub TxtFOB1_GotFocus()
   MarcarTexto TxtFOB1
End Sub

Private Sub TxtPVP1_LostFocus()
   Entrada = CCur(TxtCant1)
   ValorUnit = Val(CCur(TxtCIF1))
 ' Llenamos el ultimo saldo del kardex
   If Entrada > 0 And ValorUnit > 0 Then
      ValorTotal = ValorUnit * Entrada
      SetFields AdoKardex, "DH", "1"
      SetFields AdoKardex, "CODIGO_INV", CodigoInv
      SetFields AdoKardex, "PRODUCTO", Producto
      SetFields AdoKardex, "CANT_ES", Entrada
      SetFields AdoKardex, "VALOR_FOB", Val(CCur(TxtFOB1))
      SetFields AdoKardex, "TRANS_UNIT", Val(CCur(TxtTransUnit1))
      SetFields AdoKardex, "TRANS_TOTAL", Val(CCur(Label23.Caption))
      SetFields AdoKardex, "PRECIO_CIF", Val(CCur(TxtCIF1))
      SetFields AdoKardex, "COMIS", Val(CCur(TxtCom1)) / 100
      SetFields AdoKardex, "UTIL", Val(CCur(TxtUtil1)) / 100
      SetFields AdoKardex, "PVP", Val(CCur(TxtPVP1))
      SetFields AdoKardex, "VALOR_UNIT", ValorUnit
      SetFields AdoKardex, "VALOR_TOTAL", ValorTotal
      SetFields AdoKardex, "ORDEN", Val(CCur(TextOrden))
      SetUpdate AdoKardex
      DCInv.SetFocus
   End If
   FrmCambios.Visible = False
End Sub

Private Sub TxtTranspor_GotFocus()
   MarcarTexto TxtTranspor
End Sub

Private Sub TxtTranspor_LostFocus()
  If Val(TxtTranspor) <= 0 Then
     MsgBox "Debe Ingresar el Monto Total"
     MBFechaI.SetFocus
  End If
End Sub

Private Sub TxtTransUnit1_GotFocus()
  MarcarTexto TxtTransUnit1
End Sub

Private Sub TxtTransUnit1_LostFocus()
   Label23.Caption = Format(CCur(TxtCant1) * CCur(TxtTransUnit1), "#,##0.00")
End Sub

Private Sub TxtUtil_GotFocus()
  MarcarTexto TxtUtil
End Sub

Private Sub TxtUtil_LostFocus()
   TextoValido TxtUtil
   TextTotal = ((Val(CCur(TxtCIF)) * Val(CCur(TxtUtil))) / 100) + Val(CCur(TxtCIF))
End Sub

Private Sub TxtUtil1_GotFocus()
  MarcarTexto TxtUtil1
End Sub

Private Sub TxtUtil1_LostFocus()
  TxtPVP1 = ((Val(CCur(TxtCIF1)) * Val(CCur(TxtUtil1))) / 100) + Val(CCur(TxtCIF1))
End Sub
