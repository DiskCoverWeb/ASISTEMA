VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FSubCtas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresar Subcuentas de Proceso"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDetalle 
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
      MaxLength       =   35
      TabIndex        =   13
      Top             =   1785
      Width           =   5580
   End
   Begin MSDataListLib.DataList DLCliente 
      Bindings        =   "SubCtas.frx":0000
      DataSource      =   "AdoCliente"
      Height          =   4155
      Left            =   105
      TabIndex        =   33
      Top             =   420
      Visible         =   0   'False
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   7329
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   16711680
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
   Begin MSDataListLib.DataList DLSubCta 
      Bindings        =   "SubCtas.frx":0019
      DataSource      =   "AdoBenef"
      Height          =   1620
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "<Ctrl+S> Inserta nuevo Beneficiario"
      Top             =   420
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2858
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
      Left            =   9870
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "SubCtas.frx":0030
      Top             =   1050
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "|x|"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1380
      Left            =   105
      TabIndex        =   20
      Top             =   4620
      Visible         =   0   'False
      Width           =   11460
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
         Left            =   6090
         TabIndex        =   30
         Top             =   945
         Width           =   3585
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
         TabIndex        =   28
         Text            =   "PICHINCHA"
         Top             =   945
         Width           =   3585
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
         TabIndex        =   26
         Text            =   "ECUADOR"
         Top             =   945
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
         TabIndex        =   24
         Top             =   420
         Width           =   7785
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
         TabIndex        =   22
         ToolTipText     =   "<Alt+F2> Codigo Automático"
         Top             =   420
         Width           =   1800
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
         Left            =   6090
         TabIndex        =   29
         Top             =   735
         Width           =   3585
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
         TabIndex        =   27
         Top             =   735
         Width           =   3585
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
         TabIndex        =   25
         Top             =   735
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
         TabIndex        =   23
         Top             =   210
         Width           =   7785
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   1065
         Left            =   9765
         TabIndex        =   31
         Top             =   210
         Width           =   1590
         BackColor       =   16776960
         Caption         =   "Grabar Datos"
         PicturePosition =   327683
         Size            =   "2805;1879"
         Picture         =   "SubCtas.frx":0034
         Accelerator     =   71
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
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
         TabIndex        =   21
         Top             =   210
         Width           =   1800
      End
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
      Left            =   9030
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "SubCtas.frx":034E
      Top             =   1050
      Width           =   855
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
      Left            =   5985
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Visible         =   0   'False
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid DGSubCta 
      Bindings        =   "SubCtas.frx":0352
      Height          =   2325
      Left            =   105
      TabIndex        =   14
      Top             =   2205
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   4101
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
   Begin MSMask.MaskEdBox MBoxFechaV 
      Height          =   330
      Left            =   5985
      TabIndex        =   4
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
      Left            =   10395
      Picture         =   "SubCtas.frx":036E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4620
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
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "SubCtas.frx":07B0
      DataSource      =   "AdoFacturas"
      Height          =   315
      Left            =   7350
      TabIndex        =   7
      Top             =   1050
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
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
      TabIndex        =   12
      Top             =   1470
      Width           =   5580
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
      Left            =   9870
      TabIndex        =   10
      Top             =   735
      Width           =   1695
   End
   Begin MSForms.ToggleButton ToggleButton1 
      Height          =   540
      Left            =   5985
      TabIndex        =   32
      Top             =   105
      Visible         =   0   'False
      Width           =   750
      BackColor       =   16761024
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1323;952"
      Value           =   "0"
      Picture         =   "SubCtas.frx":07CA
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
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
      Left            =   9030
      TabIndex        =   8
      Top             =   735
      Width           =   855
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
      Left            =   7350
      TabIndex        =   6
      Top             =   735
      Width           =   1695
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
      Left            =   5985
      TabIndex        =   3
      Top             =   735
      Width           =   1380
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
      TabIndex        =   2
      Top             =   105
      Width           =   4740
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
      Left            =   7875
      TabIndex        =   19
      Top             =   5145
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
      Left            =   6300
      TabIndex        =   16
      Top             =   5145
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
      Left            =   7875
      TabIndex        =   18
      Top             =   4620
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
      Left            =   6300
      TabIndex        =   17
      Top             =   4620
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
  SelectData AdoAux, sSQL
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
  Unload FSubCtas
End Sub

Private Sub CommandButton1_Click()
  Si_No = True
  TextoValido TxtCI_RUC, , True
  TextoValido TxtApellidosS, , True
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "ORDER BY Cliente "
  SelectAdodc AdoListCtas, sSQL
  If TxtCI_RUC.Text = Ninguno Then
     MsgBox "No se puede grabar, La C.I./R.U.C. deben tener valores"
     Si_No = False
  End If
  If Si_No Then
     Mensajes = "Esta seguro de Grabar"
     Titulo = "Pregunta de Grabación"
     If BoxMensaje = vbYes Then
        Control_Procesos Normal, "Insertar Clientes en Submodulos"
        RatonReloj
        Nuevo = True
        Codigo = Ninguno
        With AdoListCtas.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("CI_RUC = '" & TxtCI_RUC & "' ")
             If .EOF Then
                 DigVerif = Digito_Verificador(TxtCI_RUC.Text)
                 Codigo = Tipo_RUC_CI.Codigo_RUC_CI
                .MoveFirst
                .Find ("Codigo = '" & Codigo & "' ")
                 If Not .EOF Then Nuevo = False
             End If
         End If
        End With
        If Nuevo Then
           SetAddNew AdoListCtas
           SetFields AdoListCtas, "T", Normal
           SetFields AdoListCtas, "Codigo", Tipo_RUC_CI.Codigo_RUC_CI
           SetFields AdoListCtas, "Cliente", TxtApellidosS
           SetFields AdoListCtas, "CI_RUC", TxtCI_RUC
           SetFields AdoListCtas, "Fecha", FechaSistema
           SetFields AdoListCtas, "Fecha_N", FechaSistema
           SetFields AdoListCtas, "Direccion", "SD"
           SetFields AdoListCtas, "DirNumero", "SN"
           SetFields AdoListCtas, "TD", Tipo_RUC_CI.Tipo_Beneficiario
           SetFields AdoListCtas, "Telefono", "022000000"
           SetFields AdoListCtas, "Pais", "593"
           SetFields AdoListCtas, "Prov", SinEspaciosIzq(CProvincia.Text)
           SetFields AdoListCtas, "Ciudad", CCiudadS
           SetFields AdoListCtas, "Grupo", NumEmpresa
           SetFields AdoListCtas, "CodigoU", CodigoUsuario
           SetUpdate AdoListCtas
           CodigoCliente = Tipo_RUC_CI.Codigo_RUC_CI
           Insertar_CxP
        End If
     End If
  End If
  sSQL = "SELECT Cl.Cliente As NomCuenta,CP.Codigo " _
       & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
       & "WHERE CP.TC = '" & SubCta & "' " _
       & "AND CP.Cta = '" & SubCtaGen & "' " _
       & "AND CP.Item = '" & NumEmpresa & "' " _
       & "AND CP.Periodo = '" & Periodo_Contable & "' " _
       & "AND Cl.Codigo <> '.' " _
       & "AND CP.Codigo = Cl.Codigo " _
       & "ORDER BY Cl.Cliente "
  SelectDBList DLSubCta, AdoBenef, sSQL, "NomCuenta"
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
  SelectData AdoAux, sSQL
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

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  Factura_No = CCur(Val(DCFactura.Text))
  TextValor = "0.00"
  With AdoFacturas.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Factura_No & " ")
       If Not .EOF Then
          If OpcTM = 1 Then
             TextValor = .Fields("Saldos_MN")
             TxtDetalle = .Fields("Detalle_SubCta")
          Else
             TextValor = .Fields("Saldos_ME")
             TxtDetalle = .Fields("Detalle_SubCta")
          End If
       End If
   End If
  End With
End Sub

Private Sub DLCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape
         DLCliente.Visible = False
         DLSubCta.SetFocus
    Case vbKeyReturn
         With AdoCliente.Recordset
          If .RecordCount > 0 Then
             .MoveFirst
             .Find ("Cliente = '" & DLCliente & "' ")
              CodigoCliente = .Fields("Codigo")
              Insertar_CxP
              sSQL = "SELECT Cl.Cliente As NomCuenta,CP.Codigo " _
                   & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
                   & "WHERE CP.TC = '" & SubCta & "' " _
                   & "AND CP.Cta = '" & SubCtaGen & "' " _
                   & "AND CP.Item = '" & NumEmpresa & "' " _
                   & "AND CP.Periodo = '" & Periodo_Contable & "' " _
                   & "AND Cl.Codigo <> '.' " _
                   & "AND CP.Codigo = Cl.Codigo " _
                   & "ORDER BY Cl.Cliente "
              SelectDBList DLSubCta, AdoBenef, sSQL, "NomCuenta"
          End If
         End With
         DLCliente.Visible = False
         DLSubCta.SetFocus
  End Select
End Sub

Private Sub DLSubCta_GotFocus()
  'SumatoriaSC = SumarSubCtas
End Sub

Private Sub DLSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyS Then
     Select Case SubCta
       Case "C", "P", "CP"
            sSQL = "SELECT * " _
                 & "FROM Clientes " _
                 & "WHERE Cliente <> '.' " _
                 & "ORDER BY Cliente "
            SelectDBList DLCliente, AdoCliente, sSQL, "Cliente"
            DLCliente.Visible = True
            DLCliente.SetFocus
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
          If Codigo = "" Then Codigo = Ninguno
          Select Case SubCta
            Case "C": sSQL = "SELECT Factura,Detalle_SubCta,SUM(Debitos-Creditos) As Saldos_MN,SUM(Parcial_ME) As Saldos_ME "
            Case "P": sSQL = "SELECT Factura,Detalle_SubCta,SUM(Creditos-Debitos) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
            Case Else: sSQL = "SELECT Factura,Detalle_SubCta,SUM(Debitos-Creditos) As Saldos_MN,-SUM(Parcial_ME) As Saldos_ME "
          End Select
          sSQL = sSQL & "FROM Trans_SubCtas " _
               & "WHERE Codigo = '" & Codigo & "' " _
               & "AND TC = '" & SubCta & "' " _
               & "AND Cta = '" & SubCtaGen & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND T = 'N' " _
               & "GROUP BY Factura,Detalle_SubCta "
          Select Case SubCta
            Case "C": sSQL = sSQL & "HAVING SUM(Debitos-Creditos) > 0 "
            Case "P": sSQL = sSQL & "HAVING SUM(Creditos-Debitos) > 0 "
            Case Else: sSQL = sSQL & "HAVING SUM(Debitos-Creditos) = 0 "
          End Select
          sSQL = sSQL & "ORDER BY Factura "
          SelectDBCombo DCFactura, AdoFacturas, sSQL, "Factura"
          If AdoFacturas.Recordset.RecordCount <= 0 Then
             DCFactura.Text = "0"
             MarcarTexto DCFactura
          End If
       End If
    End If
  End With
End Sub

Private Sub Form_Activate()
   RatonReloj
   TxtPrima.Visible = False
   If OpcTM = 2 Then Label17.Caption = "VALOR M/E" Else Label17.Caption = "VALOR M/N"
   TotalSubCta = Round(TotalSubCta, 2)
   DGSubCta.Visible = False
   LabelCta.Caption = SubCtaGen & " - " & Cuenta
   Label6.Visible = True
   If SubCta = "G" Then
      Label4.Visible = False
      MBoxFechaV.Visible = False
      Label6.Caption = "TKT"
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
   SelectData AdoAux, sSQL
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
   SelectData AdoAux, sSQL
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
   ConectarAdoExecute sSQL
   
   sSQL = "SELECT * " _
        & "FROM Asiento_SC " _
        & "WHERE TC = '" & SubCta & "' " _
        & "AND Cta = '" & SubCtaGen & "' " _
        & "AND DH = '" & OpcDH & "' " _
        & "AND TM = '" & OpcTM & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   SelectDataGrid DGSubCta, AdoSubCtaDet1, sSQL
   Select Case SubCta
     Case "G", "I", "PM"
          sSQL = "SELECT Detalle As NomCuenta,Codigo " _
               & "FROM Catalogo_SubCtas " _
               & "WHERE TC = '" & SubCta & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Codigo <> '.' " _
               & "ORDER BY Detalle "
     Case "C", "P", "CP"
          ToggleButton1.Visible = True
          sSQL = "SELECT Cl.Cliente As NomCuenta,CP.Codigo " _
               & "FROM Catalogo_CxCxP As CP,Clientes As Cl " _
               & "WHERE CP.TC = '" & SubCta & "' " _
               & "AND CP.Cta = '" & SubCtaGen & "' " _
               & "AND CP.Item = '" & NumEmpresa & "' " _
               & "AND CP.Periodo = '" & Periodo_Contable & "' " _
               & "AND Cl.Codigo <> '.' " _
               & "AND CP.Codigo = Cl.Codigo " _
               & "ORDER BY Cl.Cliente "
   End Select
   SelectDBList DLSubCta, AdoBenef, sSQL, "NomCuenta"
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
             Label3.Caption = "SUBCUENTAS DE GASTO"
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
      DLSubCta.SetFocus
      RatonNormal
   Else
      MsgBox "No existe Datos Asignados para procesar"
      Unload FSubCtas
   End If
End Sub

Private Sub Form_Load()
  CentrarForm FSubCtas
  ConectarAdodc AdoAux
  ConectarAdodc AdoBenef
  ConectarAdodc AdoCliente
  ConectarAdodc AdoListCtas
  ConectarAdodc AdoFacturas
  ConectarAdodc AdoSubCtaDet1
  FSubCtas.Height = Command1.Top + Command1.Height + 600
End Sub

Private Sub MBoxFechaV_GotFocus()
  MarcarTexto MBoxFechaV
End Sub

Private Sub MBoxFechaV_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFechaV_LostFocus()
  FechaValida MBoxFechaV, False
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
  CentrarForm FSubCtas
End Sub

Private Sub TxtApellidosS_GotFocus()
  MarcarTexto TxtApellidosS
End Sub

Private Sub TxtApellidosS_LostFocus()
   TextoValido TxtApellidosS, , True
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
          & "AND Mid(CI_RUC,1,3) = '" & NumEmpresa & "' " _
          & "AND TD = 'O' " _
          & "AND NOT Cliente IN ('.','CONSUMIDOR FINAL') " _
          & "AND NOT Codigo IN ('9999999999','8888888888','7777777777','6666666666') " _
          & "AND ISNUMERIC(CI_RUC) <> 0 " _
          & "ORDER BY CI_RUC DESC "
     SelectData AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          ContadorRUCCI = Val(Mid(.Fields("CI_RUC"), 4, Len(.Fields("CI_RUC")))) + 1
          CodigoRUC = NumEmpresa & Format(ContadorRUCCI, "00000")
      End If
     End With
     TxtCI_RUC.Text = CodigoRUC
     RatonNormal
  End If
  PresionoEnter KeyCode
End Sub

Private Sub TxtCI_RUC_LostFocus()
   sSQL = "SELECT * " _
        & "FROM Clientes " _
        & "WHERE Cliente <> '.' " _
        & "ORDER BY CI_RUC "
   SelectAdodc AdoCliente, sSQL
   
   TextoValido TxtCI_RUC, , True
   DigVerif = Digito_Verificador(TxtCI_RUC)
   If DigVerif = "-" Then
      MsgBox "RUC/CEDULA INCORRECTA"
      TxtCI_RUC.SetFocus
   End If
   With AdoCliente.Recordset
    If .RecordCount > 0 And TxtCI_RUC <> Ninguno Then
        RatonReloj
       .MoveFirst
       .Find ("CI_RUC Like '" & TxtCI_RUC & "' ")
        RatonNormal
        If Not .EOF Then
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
    End If
   End With
   Label9.Caption = "* C.I./R.U.C.   [" & Tipo_RUC_CI.Tipo_Beneficiario & "]"
   If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then
      Label9.Caption = "* R A Z O N    S O C I A L"
   Else
      Label9.Caption = "* APELLIDOS Y NOMBRES"
   End If
End Sub

Private Sub TxtDetalle_GotFocus()
  MarcarTexto TxtDetalle
End Sub

Private Sub TxtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDetalle_LostFocus()
Dim ValorDHAux As Currency
Dim Fecha_V_Prest As String
Dim Es_Fecha_Fin_Mes As Boolean
  Es_Fecha_Fin_Mes = False
  TextoValido TextValor, True
  TextoValido TxtPrima, True, , 0
  TextoValido TxtDetalle
 'MsgBox TextValor.Text
  ValorDH = CCur(TextValor)
  ValorDHAux = Redondear(ValorDH, 2)
 'If TextFactura.Text = "" Then TextFactura.Text = "0"
 'MsgBox ValorDH
  FechaValida MBoxFechaV
  Fecha_V_Prest = MBoxFechaV
  If CFechaLong(Fecha_V_Prest) = CFechaLong(UltimoDiaMes(Fecha_V_Prest)) Then Es_Fecha_Fin_Mes = True
 'MsgBox Fecha_V_Prest & vbCrLf & Es_Fecha_Fin_Mes & vbCrLf & UltimoDiaMes(Fecha_V_Prest)
  If ValorDH > 0 Then
     If OpcTM = 2 Then
        If Opcion_Mulp Then
           ValorDH = Round(ValorDH * Dolar, 2)
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
     If NoMeses <= 0 Then NoMeses = 1
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
          If SubCta = "PM" Then
            .Fields("Prima") = CLng(Val(DCFactura.Text))
            .Fields("Factura") = CLng(Val(TxtPrima))
          Else
            .Fields("Factura") = CLng(Val(DCFactura.Text))
          End If
         .Fields("Codigo") = Codigo
         .Fields("Beneficiario") = DLSubCta.Text
         .Fields("Detalle_SubCta") = TxtDetalle
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

Private Sub TxtMeses_GotFocus()
  MarcarTexto TxtMeses
End Sub

Private Sub TxtMeses_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtMeses_LostFocus()
  TextoValido TxtMeses, True, , 0
End Sub

Private Sub TxtPrima_GotFocus()
  MarcarTexto TxtPrima
End Sub

Private Sub TxtPrima_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPrima_LostFocus()
  TextoValido TxtPrima, True, , 0
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
     ConectarAdoExecute sSQL
    'Procedemos a grabar submodulo
     sSQL = "SELECT * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Cta = '" & SubCtaGen & "' " _
          & "AND Codigo = '" & CodigoCliente & "' " _
          & "AND TC = '" & SubCta & "' "
     SelectData AdoAux, sSQL
    'Grabamos al Catalogo de Sub modulos
     SetAddNew AdoAux
     SetFields AdoAux, "Item", NumEmpresa
     SetFields AdoAux, "Periodo", Periodo_Contable
     SetFields AdoAux, "Codigo", CodigoCliente
     SetFields AdoAux, "Cta", SubCtaGen
     SetFields AdoAux, "TC", SubCta
     SetFields AdoAux, "Importaciones", 0
     SetUpdate AdoAux
End Sub
