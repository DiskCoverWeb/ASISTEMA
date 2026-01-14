VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FSubCtas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "z"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SubCtas.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextValor 
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
      Left            =   12075
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "SubCtas.frx":0342
      Top             =   1050
      Width           =   1695
   End
   Begin VB.TextBox TxtMeses 
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
      Left            =   11235
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "SubCtas.frx":0346
      Top             =   1050
      Width           =   855
   End
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   9870
      TabIndex        =   12
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1050
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
   Begin MSMask.MaskEdBox MBoxFechaE 
      Height          =   330
      Left            =   8505
      TabIndex        =   10
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   1050
      Visible         =   0   'False
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
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "SubCtas.frx":034A
      DataSource      =   "AdoFacturas"
      Height          =   345
      Left            =   6825
      TabIndex        =   8
      Top             =   1050
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      Text            =   "Factura"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtSerie 
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
      Left            =   5985
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "SubCtas.frx":0364
      Top             =   1050
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "SubCtas.frx":036B
      DataSource      =   "AdoCliente"
      Height          =   2655
      Left            =   210
      TabIndex        =   41
      Top             =   2205
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      BackColor       =   12648447
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCDetalle 
      Bindings        =   "SubCtas.frx":0384
      DataSource      =   "AdoDetalle"
      Height          =   345
      Left            =   5985
      TabIndex        =   18
      Top             =   1785
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox ToggleButton1 
      Height          =   540
      Left            =   5985
      Picture         =   "SubCtas.frx":039D
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   105
      Width           =   750
   End
   Begin MSDataListLib.DataList DLSubCta 
      Bindings        =   "SubCtas.frx":07DF
      DataSource      =   "AdoBenef"
      Height          =   1410
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "<Ctrl+S> Inserta nuevo Beneficiario"
      Top             =   735
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2487
      _Version        =   393216
      MatchEntry      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "|x|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2010
      Left            =   105
      TabIndex        =   25
      Top             =   5145
      Visible         =   0   'False
      Width           =   13665
      Begin VB.CommandButton Command2 
         Caption         =   "&No Grabar"
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
         Left            =   12390
         Picture         =   "SubCtas.frx":07F6
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtEmail2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6300
         MaxLength       =   60
         TabIndex        =   32
         Top             =   945
         Width           =   6000
      End
      Begin VB.TextBox TxtEmail1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   31
         Top             =   945
         Width           =   6210
      End
      Begin VB.CommandButton CommandButton1 
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
         Left            =   12390
         Picture         =   "SubCtas.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   210
         Width           =   1170
      End
      Begin VB.ComboBox CCiudadS 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6300
         TabIndex        =   38
         Top             =   1470
         Width           =   6000
      End
      Begin VB.ComboBox CProvincia 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   36
         Text            =   "PICHINCHA"
         Top             =   1470
         Width           =   3795
      End
      Begin VB.ComboBox CNacion 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   34
         Text            =   "ECUADOR"
         Top             =   1470
         Width           =   2430
      End
      Begin VB.TextBox TxtApellidosS 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1890
         MaxLength       =   60
         TabIndex        =   29
         Top             =   420
         Width           =   10410
      End
      Begin VB.TextBox TxtCI_RUC 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   27
         ToolTipText     =   "<Alt+F2> Codigo Automático"
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CORREOS ELECTRONICOS"
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
         TabIndex        =   30
         Top             =   735
         Width           =   12195
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CIUDAD"
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
         Left            =   6300
         TabIndex        =   37
         Top             =   1260
         Width           =   6000
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVINCIA"
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
         TabIndex        =   35
         Top             =   1260
         Width           =   3795
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NACIONALIDAD"
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
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
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
         Height          =   225
         Left            =   1890
         TabIndex        =   28
         Top             =   210
         Width           =   10410
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   105
         TabIndex        =   26
         Top             =   210
         Width           =   1800
      End
   End
   Begin VB.TextBox TxtPrima 
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
      Left            =   6825
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "SubCtas.frx":198A
      Top             =   1050
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DGSubCta 
      Bindings        =   "SubCtas.frx":198F
      Height          =   2850
      Left            =   105
      TabIndex        =   19
      Top             =   2205
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   5027
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16761024
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   315
      Top             =   2310
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Continuar"
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
      Left            =   12600
      Picture         =   "SubCtas.frx":19AB
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5145
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   315
      Top             =   2625
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
      Caption         =   "Facturas"
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
   Begin MSAdodcLib.Adodc AdoSubCtaDet1 
      Height          =   330
      Left            =   315
      Top             =   2940
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
      Caption         =   "SubCtaDet1"
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
      Left            =   315
      Top             =   3255
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
      Left            =   315
      Top             =   3570
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   315
      Top             =   3885
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   315
      Top             =   4200
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
   Begin MSDataListLib.DataCombo DCSubCta 
      Bindings        =   "SubCtas.frx":1DED
      DataSource      =   "AdoNivel"
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoNivel 
      Height          =   330
      Left            =   315
      Top             =   4515
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
      Caption         =   "Nivel"
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
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALOR"
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
      TabIndex        =   15
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MESES"
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
      TabIndex        =   13
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA VEN."
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
      TabIndex        =   11
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA EMI."
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
      TabIndex        =   9
      Top             =   735
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FACTURA No."
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
      Left            =   6825
      TabIndex        =   6
      Top             =   735
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SERIE"
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
      Left            =   5985
      TabIndex        =   4
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DETALLE AUXILIAR DEL SUBMODULO"
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
      Left            =   5985
      TabIndex        =   17
      Top             =   1470
      Width           =   7785
   End
   Begin VB.Label LabelCta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE CUENTA"
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
      Height          =   540
      Left            =   6825
      TabIndex        =   3
      Top             =   105
      Width           =   6945
   End
   Begin VB.Label LabelTotalSCME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   10080
      TabIndex        =   24
      Top             =   5670
      Width           =   2430
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL M/E"
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
      Left            =   8505
      TabIndex        =   21
      Top             =   5670
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B E N E F I C I A R I O"
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
      Top             =   105
      Width           =   5790
   End
   Begin VB.Label LabelTotalSCMN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   10080
      TabIndex        =   23
      Top             =   5145
      Width           =   2430
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TOTAL M/N"
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
      Left            =   8505
      TabIndex        =   22
      Top             =   5145
      Width           =   1590
   End
End
Attribute VB_Name = "FSubCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim FilaActual As Integer
Dim SumaSubCta As Currency
Dim SumaSubCta_ME As Currency
Dim AgruparSubMod As Boolean
Dim InsSubMod As Boolean
Dim FechaEmi As String

Public Function SumarSubCtas() As Currency
  SumaSubCta = 0: SumaSubCta_ME = 0
  'DataSubCtaDet1.Refresh
  With AdoSubCtaDet1.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
         'If OpcDH = .Fields("DH") Then
          SumaSubCta = SumaSubCta + .Fields("Valor")
          SumaSubCta_ME = SumaSubCta_ME + .Fields("Valor_ME")
         'End If
         .MoveNext
     Loop
   End If
  End With
  SumaSubCta = Round(SumaSubCta, 2)
  SumaSubCta_ME = Round(SumaSubCta_ME, 2)
  LabelTotalSCMN.Caption = Format(SumaSubCta, "#,##0.00")
  LabelTotalSCME.Caption = Format(SumaSubCta_ME, "#,##0.00")
  If OpcTM = 2 Then
     SumatoriaSC = SumaSubCta_ME
     SumarSubCtas = SumaSubCta_ME
  Else
     SumatoriaSC = SumaSubCta
     SumarSubCtas = SumaSubCta
  End If
End Function

Private Sub CCiudadS_GotFocus()
  MarcarTexto CCiudadS
End Sub

Private Sub CCiudadS_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CNacion_GotFocus()
  MarcarTexto CNacion
End Sub

Private Sub CNacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CNacion_LostFocus()
  CProvincia.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'P' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "ORDER BY CProvincia "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CProvincia.Text = AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CProvincia.AddItem AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CProvincia.AddItem "99 OTRO"
     CProvincia.Text = "99 OTRO"
  End If
End Sub

Private Sub Command1_Click()
  SumatoriaSC = SumarSubCtas
  Unload FSubCtas
End Sub

Private Sub Command2_Click()
  FSubCtas.Height = Command1.Top + Command1.Height + 600
  ToggleButton1.value = False
  Frame1.Visible = False
  DLSubCta.SetFocus
End Sub

Private Sub CommandButton1_Click()
  TextoValido TxtCI_RUC, , True
  TextoValido TxtApellidosS, , True
  If TxtCI_RUC.Text = Ninguno Or TxtApellidosS = Ninguno Then
     MsgBox "No se puede grabar, La C.I./R.U.C. deben tener valores"
  Else
     Mensajes = "Esta seguro de Grabar"
     Titulo = "Pregunta de Grabación"
     If BoxMensaje = vbYes Then
        Control_Procesos Normal, "Insertar Clientes en Submodulos"
        RatonReloj
        Codigo = Ninguno
        sSQL = "SELECT Codigo " _
             & "FROM Clientes " _
             & "WHERE CI_RUC = '" & TxtCI_RUC & "' "
        Select_Adodc AdoListCtas, sSQL
        If AdoListCtas.Recordset.RecordCount <= 0 Then
           DigVerif = Digito_Verificador(TxtCI_RUC)
           Codigo = Tipo_RUC_CI.Codigo_RUC_CI
           CodigoCliente = Codigo
           sSQL = "SELECT CI_RUC " _
                & "FROM Clientes " _
                & "WHERE Codigo = '" & Codigo & "' "
           Select_Adodc AdoListCtas, sSQL
           If AdoListCtas.Recordset.RecordCount <= 0 Then
              SetAdoAddNew "Clientes"
              SetAdoFields "T", Normal
              SetAdoFields "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
              SetAdoFields "Cliente", UCaseStrg(TrimStrg(TxtApellidosS))
              SetAdoFields "CI_RUC", TxtCI_RUC
              SetAdoFields "Fecha", FechaSistema
              SetAdoFields "Fecha_N", FechaSistema
              SetAdoFields "Direccion", "SD"
              SetAdoFields "DirNumero", "SN"
              SetAdoFields "TD", Tipo_RUC_CI.Tipo_Beneficiario
              SetAdoFields "Telefono", "022000000"
              SetAdoFields "Pais", "593"
              SetAdoFields "Prov", SinEspaciosIzq(CProvincia.Text)
              SetAdoFields "Ciudad", CCiudadS
              SetAdoFields "Grupo", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoFields "Email", TxtEmail1
              SetAdoFields "Email2", TxtEmail1
              SetAdoUpdate
              Insertar_CxP
           Else
              MsgBox "No se puede volver a crear un Codigo Existente"
           End If
        Else
            MsgBox "No se puede volver a crear un CI/RUC Existente"
        End If
     End If
  End If
  sSQL = "SELECT Cl.Cliente As NomCuenta,CP.Codigo, Cl.Credito " _
       & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
       & "WHERE CP.TC = '" & SubCta & "' " _
       & "AND CP.Cta = '" & SubCtaGen & "' " _
       & "AND CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND Cl.Codigo <> '.' " _
       & "AND CP.Codigo = Cl.Codigo " _
       & "ORDER BY Cl.Cliente "
  SelectDB_List DLSubCta, AdoBenef, sSQL, "NomCuenta"
  FSubCtas.Height = Command1.Top + Command1.Height + 600
  ToggleButton1.value = False
  Frame1.Visible = False
  DLSubCta.SetFocus
End Sub

Private Sub CProvincia_GotFocus()
  MarcarTexto CProvincia
End Sub

Private Sub CProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CProvincia_LostFocus()
  CCiudadS.Clear
  sSQL = "SELECT * " _
       & "FROM Tabla_Naciones " _
       & "WHERE CProvincia <> '00' " _
       & "AND TR = 'C' " _
       & "AND CPais = '" & SinEspaciosIzq(CNacion) & "' " _
       & "AND CProvincia = '" & SinEspaciosIzq(CProvincia) & "' " _
       & "ORDER BY CCiudad "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     CCiudadS.Text = AdoAux.Recordset.Fields("Descripcion_Rubro")
     Do While Not AdoAux.Recordset.EOF
        CCiudadS.AddItem AdoAux.Recordset.Fields("Descripcion_Rubro")
        AdoAux.Recordset.MoveNext
     Loop
  Else
     CCiudadS.AddItem "OTRO"
     CCiudadS.Text = "OTRO"
  End If
End Sub

Private Sub DCDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCDetalle_LostFocus()
Dim ValorDHAux As Currency
Dim Fecha_V_Prest As String
Dim Es_Fecha_Fin_Mes As Boolean
Dim SCCantidad As Currency
  Es_Fecha_Fin_Mes = False
  TextoValido TextValor, True
  TextoValido TxtPrima, True, , 0
  TextoValido TxtMeses, True, , 0
  
 'MsgBox TextValor.Text
  ValorDH = CCur(TextValor)
  ValorDHAux = Redondear(ValorDH, 2)
  If Len(TxtSerie) < 6 Then TxtSerie = "001001"
 'If TextFactura.Text = "" Then TextFactura.Text = "0"
 'MsgBox ValorDH
  If DCDetalle.Text = "" Then DCDetalle.Text = Ninguno
  FechaValida MBoxFechaV
  Fecha_V_Prest = MBoxFechaV
  If CFechaLong(Fecha_V_Prest) = CFechaLong(UltimoDiaMes(Fecha_V_Prest)) Then Es_Fecha_Fin_Mes = True
 'MsgBox Fecha_V_Prest & vbCrLf & Es_Fecha_Fin_Mes & vbCrLf & UltimoDiaMes(Fecha_V_Prest)
  If MBoxFechaE.Visible Then FechaEmi = MBoxFechaE.Text Else FechaEmi = Co.Fecha

  If ValorDH > 0 Then
     If OpcTM = 2 Then
        If Opcion_Mulp Then
           ValorDH = Redondear(ValorDH * Dolar, 2)
        Else
           If Dolar <= 0 Then
              MsgBox "No se puede Dividir para cero," & vbCrLf & "cambie la Cotización."
              ValorDH = 0
           Else
              ValorDH = Redondear(ValorDH / Dolar, 2)
           End If
        End If
     End If
     NoMeses = Val(TxtMeses)
     SCCantidad = Val(TxtMeses)
     If NoMeses <= 0 Then NoMeses = 1
     Select Case SubCta
       Case "G", "I", "PM": NoMeses = 1
     End Select
     With AdoSubCtaDet1.Recordset
      For I = 1 To NoMeses
         .AddNew
         .Fields("Prima") = 0
          If SubCta = "G" Then
            .Fields("Fecha_V") = FechaTexto
          ElseIf SubCta = "PM" Then
            .Fields("Fecha_V") = FechaTexto
          Else
            .Fields("Fecha_V") = Fecha_V_Prest
          End If
         .Fields("TC") = SubCta
         .Fields("Serie") = TxtSerie.Text
         .Fields("FECHA_E") = FechaEmi
          If SubCta = "PM" Then
            .Fields("Prima") = CLng(Val(DCFactura.Text))
            .Fields("Factura") = CLng(Val(TxtPrima))
          ElseIf SubCta = "G" Then
            .Fields("Factura") = CLng(Val(DCFactura.Text))
            .Fields("Prima") = CCur(Val(SCCantidad))
          Else
             If NoMeses > 1 Then
               .Fields("Factura") = Val(Format(FechaTexto, "YYMMDD") & Format(I, "00"))
             Else
               .Fields("Factura") = CLng(Val(DCFactura.Text))
             End If
          End If
         .Fields("Codigo") = Codigo
         .Fields("Beneficiario") = TrimStrg(MidStrg(DLSubCta.Text, 1, 60))
         .Fields("Detalle_SubCta") = TrimStrg(MidStrg(DCDetalle.Text, 1, 60))
         .Fields("Cta") = SubCtaGen
         .Fields("DH") = OpcDH
         .Fields("Valor") = ValorDH
         .Fields("Valor_ME") = 0
         .Fields("TM") = OpcTM
         .Fields("Item") = NumEmpresa
         .Fields("T_No") = Trans_No
         .Fields("SC_No") = LnSC_No
         .Fields("CodigoU") = CodigoUsuario
          If OpcTM = 2 Then .Fields("Valor_ME") = ValorDHAux
         .Update
          Fecha_V_Prest = SiguienteMes(Fecha_V_Prest, Es_Fecha_Fin_Mes)
      Next I
      LnSC_No = LnSC_No + 1
     End With
  End If
  SumatoriaSC = SumarSubCtas
  DLSubCta.SetFocus
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  If IsNumeric(DCFactura.Text) Then Factura_No = CCur(Val(DCFactura.Text)) Else Factura_No = 0
  TextValor = "0.00"
  Select Case SubCta
    Case "C", "P"
         With AdoFacturas.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Factura = " & Factura_No & " ")
             If Not .EOF Then
                If OpcTM = 1 Then
                   TextValor = .Fields("Saldos_MN")
                Else
                   TextValor = .Fields("Saldos_ME")
                End If
                If AgruparSubMod Then TxtDetalle = .Fields("Detalle_SubCta") Else TxtDetalle = Ninguno
             End If
         End If
        End With
  End Select
    
    sSQL = "SELECT Serie, Factura, COUNT(Factura) As CantFact " _
         & "FROM Trans_SubCtas " _
         & "WHERE TC = '" & SubCta & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Serie = '" & TxtSerie.Text & "' " _
         & "AND Factura = " & DCFactura.Text & " " _
         & "GROUP BY Serie, Factura " _
         & "ORDER BY Serie, Factura "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       Label13.Visible = False
       MBoxFechaE.Visible = False
    Else
       Label13.Visible = True
       MBoxFechaE.Visible = True
       MBoxFechaE.SetFocus
    End If
End Sub

Private Sub DCSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCSubCta_LostFocus()
Dim Nivel_No As String
    Select Case SubCta
      Case "G", "I", "PM"
           Nivel_No = "00"
           With AdoNivel.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Detalle = '" & DCSubCta & "' ")
                If Not .EOF Then Nivel_No = .Fields("Nivel")
            End If
           End With
           sSQL = "SELECT Detalle As NomCuenta,Codigo, 0 As Credito " _
                & "FROM Catalogo_SubCtas " _
                & "WHERE TC = '" & SubCta & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Agrupacion = " & Val(adFalse) & " " _
                & "AND Nivel = '" & Nivel_No & "' " _
                & "AND Codigo <> '.' " _
                & "ORDER BY Nivel,Detalle "
           SelectDB_List DLSubCta, AdoBenef, sSQL, "NomCuenta"
    End Select
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         InsSubMod = False
         DCCliente.Visible = False
         DLSubCta.SetFocus
    Case vbKeyReturn
         SiguienteControl
  End Select
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Codigo, CI_RUC, Cliente " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' " Else sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoCliente, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
    With AdoCliente.Recordset
     If .RecordCount > 0 And InsSubMod Then
        .MoveFirst
        .Find ("Cliente = '" & DCCliente.Text & "' ")
         If Not .EOF Then CodigoCliente = .Fields("Codigo")
         Insertar_CxP
         Listar_SubCta_Modulo
     End If
    End With
    DCCliente.Visible = False
    DLSubCta.SetFocus
End Sub

Private Sub DLSubCta_GotFocus()
  'SumatoriaSC = SumarSubCtas
End Sub

Private Sub DLSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyS Then
     InsSubMod = True
     Select Case SubCta
       Case "C", "P", "CP"
            sSQL = "SELECT TOP 50 Codigo, Cliente, CI_RUC, Credito " _
                 & "FROM Clientes " _
                 & "WHERE LEN(Cliente) > 1 " _
                 & "ORDER BY Cliente "
            SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
            DCCliente.Visible = True
            DCCliente.SetFocus
     End Select
  End If
  If KeyCode = vbKeyEscape Then Command1.SetFocus
  PresionoEnter KeyCode
End Sub

Private Sub DLSubCta_LostFocus()
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("NomCuenta = '" & DLSubCta.Text & "' ")
       If Not .EOF Then
          Codigo = .Fields("Codigo")
         'MsgBox .Fields("Credito") & " - " & Co.Fecha
          If MBoxFechaV.Visible Then MBoxFechaV = CLongFecha(CFechaLong(Co.Fecha) + .Fields("Credito"))
          If Codigo = "" Then Codigo = Ninguno
       End If
    End If
  End With
End Sub

Private Sub Form_Activate()
   RatonReloj
   AgruparSubMod = Datos_De_Empresa("Det_SubMod")
   DCSubCta.Visible = False
   PorCtasCostos = False
   If MidStrg(SubCtaGen, 1, 1) = "1" Then
      Select Case SubCta
        Case "G", "CC"
             sSQL = "SELECT Cta " _
                  & "FROM Trans_Presupuestos " _
                  & "WHERE Periodo = '" & Periodo_Contable & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND MesNo = 0 " _
                  & "GROUP BY Cta "
             Select_Adodc AdoAux, sSQL
             If AdoAux.Recordset.RecordCount > 0 Then PorCtasCostos = True
      End Select
   End If
   MBoxFechaV = FechaTexto
   MBoxFechaE = FechaTexto
   TxtPrima.Visible = False
   If OpcTM = 2 Then Label17.Caption = "VALOR M/E" Else Label17.Caption = "VALOR M/N"
   TotalSubCta = Redondear(TotalSubCta, 2)
   DGSubCta.Visible = False
   LabelCta.Caption = SubCtaGen & " - " & Cuenta
   Label6.Visible = True
   Label2.Caption = "MESES"
   If SubCta = "G" Then
      Label4.Visible = False
      MBoxFechaV.Visible = False
      Label6.Caption = "VALOR"
      Label2.Caption = "CANT."
   ElseIf SubCta = "PM" Then
      Label4.Visible = True
      MBoxFechaV.Visible = False
      Label4.Caption = "FACTURA No."
      TxtPrima.Visible = True
      Label6.Caption = "PRIMA"
   Else
      Label4.Visible = True
      MBoxFechaV.Visible = True
      Label6.Caption = "Factura No."
   End If
   DGSubCta.Visible = True
   SumaSubCta = 0: SumaSubCta_ME = 0
   LabelTotalSCMN.Caption = Format(0, "#,##0.00")
   LabelTotalSCME.Caption = Format(0, "#,##0.00")
   
   CNacion.Clear
   sSQL = "SELECT * " _
        & "FROM Tabla_Naciones " _
        & "WHERE TR = 'N' " _
        & "ORDER BY CPais,Descripcion_Rubro "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      CNacion.Text = "593 ECUADOR"
      Do While Not AdoAux.Recordset.EOF
         CNacion.AddItem AdoAux.Recordset.Fields("CPais") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
         AdoAux.Recordset.MoveNext
      Loop
   End If
   CNacion.AddItem "999 OTRO"
   CProvincia.Clear
   sSQL = "SELECT * " _
        & "FROM Tabla_Naciones " _
        & "WHERE CProvincia <> '00' " _
        & "AND TR = 'P' " _
        & "ORDER BY CProvincia "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      CProvincia.Text = AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
      Do While Not AdoAux.Recordset.EOF
         CProvincia.AddItem AdoAux.Recordset.Fields("CProvincia") & " " & AdoAux.Recordset.Fields("Descripcion_Rubro")
         AdoAux.Recordset.MoveNext
      Loop
   End If
            
   sSQL = "DELETE * " _
        & "FROM Asiento_SC " _
        & "WHERE TC = '" & SubCta & "' " _
        & "AND Cta = '" & SubCtaGen & "' " _
        & "AND DH = '" & OpcDH & "' " _
        & "AND TM = '" & OpcTM & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Ejecutar_SQL_SP sSQL
   
   sSQL = "SELECT * " _
        & "FROM Asiento_SC " _
        & "WHERE TC = '" & SubCta & "' " _
        & "AND Cta = '" & SubCtaGen & "' " _
        & "AND DH = '" & OpcDH & "' " _
        & "AND TM = '" & OpcTM & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   Select_Adodc_Grid DGSubCta, AdoSubCtaDet1, sSQL
   
   If PorCtasCostos Then
      sSQL = "SELECT Detalle,Codigo, Nivel " _
           & "FROM Catalogo_SubCtas " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND 1 = 0 "
   Else
      sSQL = "SELECT Detalle,Codigo, Nivel " _
           & "FROM Catalogo_SubCtas " _
           & "WHERE TC = '" & SubCta & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Agrupacion <> " & Val(adFalse) & " " _
           & "AND Codigo <> '.' " _
           & "ORDER BY Nivel,Detalle "
   End If
   SelectDB_Combo DCSubCta, AdoNivel, sSQL, "Detalle"
   If AdoNivel.Recordset.RecordCount > 0 Then DCSubCta.Visible = True
   
   sSQL = "SELECT Detalle_SubCta " _
        & "FROM Trans_SubCtas " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND TC = '" & SubCta & "' " _
        & "GROUP BY Detalle_SubCta " _
        & "ORDER BY Detalle_SubCta "
   SelectDB_Combo DCDetalle, AdoDetalle, sSQL, "Detalle_SubCta"
   
   Listar_SubCta_Modulo
   RatonNormal
   If AdoBenef.Recordset.RecordCount > 0 Then
      Select Case SubCta
        Case "C"
             FSubCtas.Caption = "Ingreso se Subcuenta por Cobrar"
             Label3.Caption = "SUBCUENTAS POR COBRAR"
        Case "P"
             FSubCtas.Caption = "Ingreso se Subcuenta por Pagar"
             Label3.Caption = "SUBCUENTAS POR PAGAR"
        Case "G"
             FSubCtas.Caption = "Ingreso se Subcuenta de Gastos"
             Label3.Caption = "SUBCUENTAS DE GASTOS"
        Case "I"
             FSubCtas.Caption = "Ingreso se Subcuenta de Ingreso"
             Label3.Caption = "SUBCUENTAS DE INGRESO"
        Case "CP"
             FSubCtas.Caption = "Ingreso se Subcuenta por Cobrar"
             Label3.Caption = "SUBCUENTAS POR COBRAR PRESTAMOS"
        Case "PM"
             FSubCtas.Caption = "Ingreso se Subcuenta de Ingreso"
             Label3.Caption = "SUBCUENTAS DE PRIMAS"
        Case Else: Unload FSubCtas
      End Select
      
     'MsgBox DCSubCta.Visible
      Select Case SubCta
        Case "G", "I", "PM": If DCSubCta.Visible Then DCSubCta.SetFocus Else DLSubCta.SetFocus
        Case "C", "P", "CP": DLSubCta.SetFocus
      End Select
   Else
      MsgBox "No existe Datos Asignados para procesar"
      Unload FSubCtas
   End If
End Sub

Private Sub Form_Load()
  CentrarForm FSubCtas
  ConectarAdodc AdoAux
  ConectarAdodc AdoBenef
  ConectarAdodc AdoNivel
  ConectarAdodc AdoCliente
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoListCtas
  ConectarAdodc AdoFacturas
  ConectarAdodc AdoSubCtaDet1
  FSubCtas.Height = Command1.Top + Command1.Height + 600
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
  FechaTexto = MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV, False
End Sub

Private Sub MBoxFechaE_GotFocus()
  MarcarTexto MBoxFechaE
End Sub

Private Sub MBoxFechaE_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaE_LostFocus()
  FechaValida MBoxFechaE, False
End Sub

Private Sub TextValor_GotFocus()
  MarcarTexto TextValor
End Sub

Private Sub TextValor_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextValor_LostFocus()
  TextoValido TextValor, True
End Sub

Private Sub ToggleButton1_Click()
  If ToggleButton1.value Then
     FSubCtas.Height = Frame1.Top + Frame1.Height + 600
     Frame1.Caption = "| CUENTA: " & SubCtaGen & " - " & Cuenta & " |"
     Frame1.Visible = True
     TxtCI_RUC.SetFocus
  Else
     FSubCtas.Height = Command1.Top + Command1.Height + 600
     Frame1.Visible = False
  End If
  'CentrarForm FSubCtas
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtApellidosS_LostFocus()
   TextoValido TxtApellidosS, , True
End Sub

Private Sub TxtEmail1_GotFocus()
  MarcarTexto TxtEmail1
End Sub

Private Sub TxtEmail1_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtEmail1_LostFocus()
   TextoValido TxtEmail1
   TxtEmail1 = LCase(TxtEmail1)
End Sub

Private Sub TxtEmail2_GotFocus()
  MarcarTexto TxtEmail2
End Sub

Private Sub TxtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtEmail2_LostFocus()
   TextoValido TxtEmail2
   TxtEmail1 = LCase(TxtEmail2)
End Sub

Private Sub TxtCI_RUC_GotFocus()
  MarcarTexto TxtCI_RUC
End Sub

Private Sub TxtCI_RUC_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CodigoRUC As String
  Keys_Especiales Shift
  If AltDown And KeyCode = vbKeyF2 Then
     RatonReloj
     ContadorRUCCI = 1
     CodigoRUC = NumEmpresa & Format(ContadorRUCCI, "00000")
     sSQL = "SELECT CI_RUC " _
          & "FROM Clientes " _
          & "WHERE LEN(CI_RUC) <= 9 " _
          & "AND MidStrg(CI_RUC,1,3) = '" & NumEmpresa & "' " _
          & "AND TD = 'O' " _
          & "AND NOT Cliente IN ('.','CONSUMIDOR FINAL') " _
          & "AND NOT Codigo IN ('9999999999','8888888888','7777777777','6666666666') " _
          & "AND ISNUMERIC(CI_RUC) <> 0 " _
          & "ORDER BY CI_RUC DESC "
     Select_Adodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          ContadorRUCCI = Val(MidStrg(.Fields("CI_RUC"), 4, Len(.Fields("CI_RUC")))) + 1
          CodigoRUC = NumEmpresa & Format(ContadorRUCCI, "00000")
      End If
     End With
     TxtCI_RUC.Text = CodigoRUC
     RatonNormal
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtCI_RUC_LostFocus()
   TextoValido TxtCI_RUC, , True
   DigVerif = Digito_Verificador(TxtCI_RUC)
   If DigVerif = "-" Then
      MsgBox "RUC/CEDULA INCORRECTA"
      TxtCI_RUC.SetFocus
   End If
   If TxtCI_RUC <> Ninguno Then
      sSQL = "SELECT Codigo, Cliente, CI_RUC, TD " _
           & "FROM Clientes " _
           & "WHERE CI_RUC Like '" & TxtCI_RUC & "' " _
           & "ORDER BY CI_RUC "
      Select_Adodc AdoCliente, sSQL
      With AdoCliente.Recordset
       If .RecordCount > 0 Then
           If .Fields("Cliente") <> TxtApellidosS Then
               MsgBox "Este R.U.C./C.I., está asignado a " & vbCrLf & vbCrLf & .Fields("Cliente")
               FSubCtas.Height = Command1.Top + Command1.Height + 600
               ToggleButton1.value = False
               Frame1.Visible = False
               DLSubCta.SetFocus
               'TxtCI_RUC.SetFocus
           Else
               TipoBenef = .Fields("TD")
           End If
       End If
      End With
   End If
   Label9.Caption = "* C.I./R.U.C.   [" & Tipo_RUC_CI.Tipo_Beneficiario & "]"
   If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
      Label9.Caption = "* R A Z O N    S O C I A L"
   Else
      Label9.Caption = "* APELLIDOS Y NOMBRES"
   End If
End Sub

Private Sub TxtMeses_GotFocus()
  MarcarTexto TxtMeses
End Sub

Private Sub TxtMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMeses_LostFocus()
  TextoValido TxtMeses, True, , 0
  Select Case SubCta
    Case "G", "I", "PM":  TextValor = Format(Val(DCFactura) * Val(TxtMeses), "#,##0.00")
  End Select
End Sub

Private Sub TxtPrima_GotFocus()
  MarcarTexto TxtPrima
End Sub

Private Sub TxtPrima_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPrima_LostFocus()
  TextoValido TxtPrima, True, , 0
  Facturas_Pendientes_SC
End Sub

Public Sub Insertar_CxP()
    'Garantizamos que no exista duplicidad
     sSQL = "DELETE * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cta = '" & SubCtaGen & "' " _
          & "AND Codigo = '" & CodigoCliente & "' " _
          & "AND TC = '" & SubCta & "' "
     Ejecutar_SQL_SP sSQL
    'Procedemos a grabar submodulo
     SetAdoAddNew "Catalogo_CxCxP"
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Periodo", Periodo_Contable
     SetAdoFields "Codigo", CodigoCliente
     SetAdoFields "Cta", SubCtaGen
     SetAdoFields "TC", SubCta
     SetAdoFields "Importaciones", 0
     SetAdoUpdate
    
     sSQL = "SELECT * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cta = '" & SubCtaGen & "' " _
          & "AND Codigo = '" & CodigoCliente & "' " _
          & "AND TC = '" & SubCta & "' "
     Select_Adodc AdoAux, sSQL
End Sub

Public Sub Facturas_Pendientes_SC()
    FechaTexto = MBoxFechaV
    If Not IsDate(FechaTexto) Then FechaTexto = FechaSistema
    If AgruparSubMod Then
       Select Case SubCta
         Case "C": sSQL = "SELECT Factura,Detalle_SubCta,(SUM(Debitos)-SUM(Creditos)) As Saldos_MN,SUM(Parcial_ME) As Saldos_ME "
         Case "P": sSQL = "SELECT Factura,Detalle_SubCta,(SUM(Creditos)-SUM(Debitos)) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
         Case Else: sSQL = "SELECT Factura,Detalle_SubCta,(SUM(Debitos)-SUM(Creditos)) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
       End Select
    Else
       Select Case SubCta
         Case "C": sSQL = "SELECT Factura,(SUM(Debitos)-SUM(Creditos)) As Saldos_MN,SUM(Parcial_ME) As Saldos_ME "
         Case "P": sSQL = "SELECT Factura,(SUM(Creditos)-SUM(Debitos)) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
         Case Else: sSQL = "SELECT Factura,(SUM(Debitos)-SUM(Creditos)) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
       End Select
    End If
    sSQL = sSQL & "FROM Trans_SubCtas " _
         & "WHERE Codigo = '" & Codigo & "' " _
         & "AND TC = '" & SubCta & "' " _
         & "AND Cta = '" & SubCtaGen & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Serie = '" & TxtSerie & "' " _
         & "AND Fecha <= #" & BuscarFecha(FechaTexto) & "# " _
         & "AND T <> 'A' "
    If AgruparSubMod Then
       sSQL = sSQL & "GROUP BY Factura,Detalle_SubCta "
    Else
       sSQL = sSQL & "GROUP BY Factura "
    End If
    Select Case SubCta
      Case "C": sSQL = sSQL & "HAVING SUM(Debitos)-SUM(Creditos) > 0 "
      Case "P": sSQL = sSQL & "HAVING SUM(Creditos)-SUM(Debitos) > 0 "
      Case Else: sSQL = sSQL & "HAVING SUM(Debitos)-SUM(Creditos) = 0 "
    End Select
    sSQL = sSQL & "ORDER BY Factura "
    SelectDB_Combo DCFactura, AdoFacturas, sSQL, "Factura"
    If AdoFacturas.Recordset.RecordCount <= 0 Then
       DCFactura.Text = "0"
       MarcarTexto DCFactura
    End If
End Sub

Public Sub Listar_SubCta_Modulo()
  'MsgBox PorCtasCostos
   Select Case SubCta
     Case "G", "I", "PM", "CC"
          If PorCtasCostos Then
             sSQL = "SELECT CS.Detalle As NomCuenta,CS.Codigo, 0 As Credito " _
                  & "FROM Catalogo_SubCtas As CS, Trans_Presupuestos As TP " _
                  & "WHERE CS.Periodo = '" & Periodo_Contable & "' " _
                  & "AND CS.Item = '" & NumEmpresa & "' " _
                  & "AND TP.Cta = '" & SubCtaGen & "' " _
                  & "AND CS.TC = '" & SubCta & "' " _
                  & "AND TP.MesNo = 0 " _
                  & "AND CS.Periodo = TP.Periodo " _
                  & "AND CS.Item = TP.Item " _
                  & "AND CS.Codigo = TP.Codigo " _
                  & "ORDER BY CS.Detalle "
          Else
             DCSubCta.Visible = True
             sSQL = "SELECT Detalle As NomCuenta,Codigo, 0 As Credito " _
                  & "FROM Catalogo_SubCtas " _
                  & "WHERE TC = '" & SubCta & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Agrupacion = " & Val(adFalse) & " " _
                  & "AND Codigo <> '.' " _
                  & "ORDER BY Nivel, Detalle "
          End If
     Case "C", "P", "CP"
          ToggleButton1.Visible = True
          sSQL = "SELECT Cl.Cliente As NomCuenta, CP.Codigo, Cl.Credito " _
               & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
               & "WHERE CP.TC = '" & SubCta & "' " _
               & "AND CP.Cta = '" & SubCtaGen & "' " _
               & "AND CP.Item = '" & NumEmpresa & "' " _
               & "AND CP.Periodo = '" & Periodo_Contable & "' " _
               & "AND Cl.Codigo <> '.' " _
               & "AND CP.Codigo = Cl.Codigo " _
               & "ORDER BY Cl.Cliente "
   End Select
   SelectDB_List DLSubCta, AdoBenef, sSQL, "NomCuenta"
End Sub

Private Sub TxtSerie_GotFocus()
  MarcarTexto TxtSerie
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtSerie_LostFocus()
    If Len(TxtSerie) < 6 Then TxtSerie = "001001"
    Facturas_Pendientes_SC
End Sub
