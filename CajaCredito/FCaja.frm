VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11145
   Begin VB.TextBox TxtPapeleta 
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
      Height          =   345
      Left            =   9345
      MaxLength       =   8
      TabIndex        =   31
      Text            =   "0"
      Top             =   5880
      Width           =   1695
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
      Left            =   9975
      Picture         =   "FCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7455
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cer&tific."
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
      Picture         =   "FCaja.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7455
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "NOMINA DE ALUMNOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4095
      TabIndex        =   45
      Top             =   420
      Visible         =   0   'False
      Width           =   6945
      Begin MSDataListLib.DataList DLAlumno 
         Bindings        =   "FCaja.frx":0BD4
         DataSource      =   "AdoAlumno"
         Height          =   2205
         Left            =   105
         TabIndex        =   46
         Top             =   210
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3889
         _Version        =   393216
         BackColor       =   8421504
         ForeColor       =   12648384
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
   Begin VB.TextBox TxtRUC 
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
      Left            =   1365
      MaxLength       =   15
      TabIndex        =   1
      Top             =   105
      Width           =   1905
   End
   Begin MSDataListLib.DataList DLTP 
      Bindings        =   "FCaja.frx":0BEC
      DataSource      =   "AdoTP"
      Height          =   2790
      Left            =   105
      TabIndex        =   17
      Top             =   5565
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4921
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "FCaja.frx":0C00
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   3990
      TabIndex        =   33
      Top             =   6615
      Visible         =   0   'False
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSDataListLib.DataList DLCtas 
      Bindings        =   "FCaja.frx":0C17
      DataSource      =   "AdoListCtas"
      Height          =   1230
      Left            =   105
      TabIndex        =   5
      Top             =   735
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2170
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
   Begin MSAdodcLib.Adodc AdoCuentas 
      Height          =   330
      Left            =   210
      Top             =   945
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
   Begin MSAdodcLib.Adodc AdoCtaNo 
      Height          =   330
      Left            =   2205
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
      Left            =   9345
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "FCaja.frx":0C31
      Top             =   5460
      Width           =   1695
   End
   Begin VB.PictureBox PicFirma 
      BackColor       =   &H00FFFFFF&
      Height          =   1800
      Left            =   105
      ScaleHeight     =   1740
      ScaleWidth      =   7410
      TabIndex        =   12
      Top             =   3360
      Width           =   7470
   End
   Begin VB.TextBox TxtNombres 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4095
      MaxLength       =   25
      TabIndex        =   3
      Top             =   105
      Width           =   6945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
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
      Left            =   7875
      Picture         =   "FCaja.frx":0C38
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7455
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Cuenta"
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
      Left            =   3990
      TabIndex        =   13
      Top             =   5250
      Width           =   3585
      Begin VB.OptionButton OpcME 
         Caption         =   "Extranjera"
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
         Left            =   1785
         TabIndex        =   15
         Top             =   210
         Width           =   1275
      End
      Begin VB.OptionButton OpcMN 
         Caption         =   "Nacional"
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
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Value           =   -1  'True
         Width           =   1170
      End
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
      Left            =   5355
      MaxLength       =   8
      TabIndex        =   40
      ToolTipText     =   "<Ctrl>+<L>: Renovación o cambio de Cartilla"
      Top             =   7875
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
      Left            =   5355
      MaxLength       =   8
      TabIndex        =   38
      Top             =   7455
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox TxtBanco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3990
      MaxLength       =   20
      TabIndex        =   34
      Top             =   6615
      Visible         =   0   'False
      Width           =   7050
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
      Left            =   6825
      Picture         =   "FCaja.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7455
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxCuenta 
      Height          =   435
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   2415
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "CCCCCCCCC-C"
      Mask            =   "########-#"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   210
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
   Begin MSAdodcLib.Adodc AdoTP 
      Height          =   330
      Left            =   105
      Top             =   6510
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
   Begin MSAdodcLib.Adodc AdoListCtas 
      Height          =   330
      Left            =   2205
      Top             =   945
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   2205
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
   Begin MSAdodcLib.Adodc AdoAlumno 
      Height          =   330
      Left            =   210
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
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "FCaja.frx":1944
      DataSource      =   "AdoClientes"
      Height          =   315
      Left            =   5355
      TabIndex        =   36
      Top             =   7035
      Visible         =   0   'False
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   4200
      Top             =   945
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4200
      Top             =   1320
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Papeleta No."
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
      Left            =   7665
      TabIndex        =   30
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label LblTipo 
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
      TabIndex        =   51
      Top             =   2835
      Width           =   10935
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fondo Reser."
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
      Left            =   7665
      TabIndex        =   50
      Top             =   3990
      Width           =   1695
   End
   Begin VB.Label LabelFondoR 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9345
      TabIndex        =   49
      Top             =   3990
      Width           =   1695
   End
   Begin VB.Label LblCartilla 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado de Cta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6825
      TabIndex        =   48
      Top             =   2100
      Width           =   2535
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paguese a:"
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
      TabIndex        =   35
      Top             =   7035
      Visible         =   0   'False
      Width           =   1380
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
      Left            =   9345
      TabIndex        =   11
      Top             =   2415
      Width           =   1695
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2415
      Width           =   7680
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
      Left            =   7665
      TabIndex        =   28
      Top             =   5460
      Width           =   1695
   End
   Begin VB.Label LabelDispRet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   9345
      TabIndex        =   27
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Disp. a Retirar"
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
      Left            =   7665
      TabIndex        =   26
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label LabelEncaje 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9345
      TabIndex        =   25
      Top             =   4620
      Width           =   1695
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Encaje"
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
      Left            =   7665
      TabIndex        =   24
      Top             =   4620
      Width           =   1695
   End
   Begin VB.Label LabelPorConf 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   9345
      TabIndex        =   23
      Top             =   4305
      Width           =   1695
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Total"
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
      Left            =   7665
      TabIndex        =   22
      Top             =   4305
      Width           =   1695
   End
   Begin VB.Label LabelEnCheques 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9345
      TabIndex        =   21
      Top             =   3675
      Width           =   1695
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Por Efectivizar"
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
      Left            =   7665
      TabIndex        =   20
      Top             =   3675
      Width           =   1695
   End
   Begin VB.Label LabelDisponible 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   9345
      TabIndex        =   19
      Top             =   3360
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
      Left            =   7665
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
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
      Left            =   9345
      TabIndex        =   10
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CUENTAS ABIERTAS"
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
      TabIndex        =   4
      Top             =   525
      Width           =   10935
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
      Left            =   1680
      TabIndex        =   8
      Top             =   2100
      Width           =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombres"
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
      Left            =   3255
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
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
      Left            =   1890
      TabIndex        =   44
      Top             =   2100
      Width           =   1590
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
      TabIndex        =   39
      Top             =   7875
      Width           =   1380
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Cuenta No."
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
      Top             =   2100
      Width           =   1590
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
      TabIndex        =   37
      Top             =   7455
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Banco"
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
      TabIndex        =   32
      Top             =   6300
      Visible         =   0   'False
      Width           =   7050
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
      TabIndex        =   16
      Top             =   5250
      Width           =   3795
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.I. / R.U.C."
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
      Width           =   1275
   End
End
Attribute VB_Name = "FCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Listar_Alumnos()
  sSQL = "SELECT C.Cliente,F.Saldo_MN,C.Codigo,F.Factura,F.TC,C.Direccion " _
       & "FROM Clientes As C,Facturas As F " _
       & "WHERE C.Codigo = F.CodigoC " _
       & "AND NOT F.TC IN ('C','P') " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.T = 'P' " _
       & "AND C.Casilla = '" & MBoxCuenta.Text & "' " _
       & "ORDER BY C.Cliente,Factura "
  SelectDBList DLAlumno, AdoAlumno, sSQL, "Cliente"
End Sub

Public Sub Insertar_Montos(DtaCta As Adodc, _
                           DtaBanc As Adodc, _
                           CuentaNo As MaskEdBox, _
                           TDebe As Currency, _
                           THaber As Currency, _
                           NoCheque As String, _
                           NomBanco As String, _
                           SaldoAnt As Currency)
Dim Saldo_Final As Currency

  If CuentaNo.Text <> "00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
 'Insertar Transacciones de Libreta
  sSQL = "SELECT TOP 1 * " _
       & "FROM Trans_Libretas " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           ID_Trans = .Fields("IDT")
       Else
           SaldoCont = SaldoAnt
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Cuenta_No") = CuentaNo.Text
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
       If TipoGrupo Then
          If THaber <> 0 Then
            '.Fields("Saldo_Disp") = SaldoDisp
            .Fields("Saldo_Disp") = SaldoCont + THaber - TDebe
            .Fields("T") = Normal
            Saldo_Final = SaldoCont + THaber - TDebe
          Else
            .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
            .Fields("T") = Normal
            Saldo_Final = SaldoDisp + THaber - TDebe
          End If
       Else
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
         .Fields("T") = Normal
          Saldo_Final = SaldoDisp + THaber - TDebe
       End If
      .Fields("CodigoU") = CodigoUsuario
       If NumeroLineas >= 36 Then NumeroLineas = 1
      .Fields("IP") = adFalse
      .Fields("CHT") = adFalse
       Select Case TipoProc
         Case "DEPC", "DEAC", "APEC"
             .Fields("CHT") = True
             .Fields("Banco") = NomBanco
             .Fields("Dias") = 2           ' Efectivizacion de cheques
             .Fields("Cheque") = NoCheque
       End Select
      .Fields("ACL") = adFalse
      .Fields("AC") = adFalse              ' Quitar
      .Fields("ACC") = adFalse
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = adFalse
      .Fields("Cartilla_No") = Cartilla_No
      .Fields("Papeleta_No") = TxtPapeleta
       SetUpdate DtaCta
       
       Select Case TipoProc
         Case "DEPC", "DEAC", "APEC", "NCCA"
              SetAdoAddNew "Trans_Bloqueos"
              SetAdoFields "T", Normal
              SetAdoFields "Fecha", FechaSistema
              SetAdoFields "Cuenta_No", CuentaNo.Text
              SetAdoFields "Valor", THaber
              SetAdoFields "Cheque", NoCheque
              If TipoProc = "NCCA" Then
                 SetAdoFields "Dias", 0
              Else
                 SetAdoFields "Dias", 2
              End If
              SetAdoFields "Banco", NomBanco
              SetAdoFields "Item", NumEmpresa
              SetAdoUpdate
       End Select
      '.Update
  End With
  If TipoProc = "CIER" Then
     Msg = UCase(InputBox("Motivo de la anulacion: ", "ACTUALIZACION DE DATOS", ""))
     Control_Procesos Normal, Msg
     sSQL = "UPDATE Clientes_Datos_Extras " _
          & "SET T = '" & Anulado & "' " _
          & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
          & "AND Tipo_Dato = 'LIBRETAS' "
     ConectarAdoExecute sSQL
  Else
     If Saldo_Final <= 0 Then
        Msg = "Cierre automático por insuficiencia de fondos"
        Control_Procesos Normal, Msg
        sSQL = "UPDATE Clientes_Datos_Extras " _
             & "SET T = '" & Anulado & "' " _
             & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
             & "AND Tipo_Dato = 'LIBRETAS' "
        ConectarAdoExecute sSQL
     Else
        sSQL = "UPDATE Clientes_Datos_Extras " _
             & "SET T = '" & Normal & "' " _
             & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
             & "AND Tipo_Dato = 'LIBRETAS' "
        ConectarAdoExecute sSQL
     End If
  End If
  End If
End Sub

Public Sub Insertar_Certif(DtaCta As Adodc, _
                           DtaBanc As Adodc, _
                           CuentaNo As MaskEdBox, _
                           TDebe As Currency, _
                           THaber As Currency, _
                           NoCheque As String, _
                           NomBanco As String, _
                           SaldoAnt As Currency)
  If CuentaNo.Text <> "00000000-0" Then
  SaldoDisp = 0: SaldoCont = 0: ID_Trans = 0
  TiempoTexto = Format(Time, FormatoTimes)
  If NumeroLineas <= 0 Then NumeroLineas = 1
 'Insertar Transacciones de Libreta
  sSQL = "SELECT TOP 1 * " _
       & "FROM Trans_Certificados " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
       & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
  SelectAdodc DtaCta, sSQL
  With DtaCta.Recordset
       If .RecordCount > 0 Then
           SaldoDisp = .Fields("Saldo_Disp")
           SaldoCont = .Fields("Saldo_Cont")
           ID_Trans = .Fields("IDT")
       Else
           SaldoCont = SaldoAnt
       End If
      .AddNew
      .Fields("Fecha") = FechaSistema
      .Fields("Cuenta_No") = CuentaNo.Text
      .Fields("TP") = TipoProc
      .Fields("Debitos") = TDebe
      .Fields("Creditos") = THaber
      .Fields("Saldo_Cont") = SaldoCont + THaber - TDebe
       If TipoGrupo Then
          If THaber <> 0 Then
            .Fields("Saldo_Disp") = SaldoDisp
            .Fields("T") = Pendiente
          Else
            .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
            .Fields("T") = Normal
          End If
       Else
         .Fields("Saldo_Disp") = SaldoDisp + THaber - TDebe
         .Fields("T") = Normal
       End If
      .Fields("CodigoU") = CodigoUsuario
       If NumeroLineas >= 36 Then NumeroLineas = 1
      .Fields("IP") = adFalse
      .Fields("CHT") = adFalse
      .Fields("ACL") = adFalse
      .Fields("ACC") = adFalse
      .Fields("IDT") = ID_Trans + 1
      .Fields("Hora") = TiempoTexto
      .Fields("Item") = NumEmpresa
      .Fields("ME") = adFalse
      .Update
  End With
  End If
End Sub

Public Sub Listar_Cuentas(CuentaNo As MaskEdBox)
Dim Total_Depositos As Currency
Dim Total_Fondo_Reserva As Currency
  Mi_Cta = False
  SaldoDisp = 0
  SaldoCont = 0
  TotalEncaje = 0
  NumeroLineas = 0
  Moneda_US = False
  LabelSocio.Caption = ""
  T = Normal
  TextLinea.Text = NumeroLineas
  LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
  LabelPorConf.Caption = Format(SaldoCont, "#,##0.00")
  LabelEnCheques.Caption = Format(SaldoCont - SaldoDisp, "#,##0.00")
  Total_Depositos = 0
  Total_Fondo_Reserva = 0
  sSQL = "SELECT * " _
       & "FROM Trans_Bloqueos " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' " _
       & "AND T = 'N' "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If .Fields("Dias") <= 0 Then
              If .Fields("Cheque") = "FR" Then
                  Total_Fondo_Reserva = Total_Fondo_Reserva + .Fields("Valor")
              Else
                  TotalEncaje = TotalEncaje + .Fields("Valor")
              End If
          Else
              Total_Depositos = Total_Depositos + .Fields("Valor")
          End If
         .MoveNext
       Loop
   End If
  End With
  CodigoCliente = "-1"
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
  If .RecordCount > 0 Then
      CodigoCliente = .Fields("Codigo")
      CodigoCli = CodigoCliente
      Moneda_US = False   '.Fields("ME")
      Mi_Cta = True
      ConLibreta = True
      T = .Fields("T")
      NoCert = .Fields("No_Soc")
      Fecha_Retiro = .Fields("Fecha_Ret")
      TipoCta = .Fields("Tipo")
      TipoDoc = .Fields("Acreditacion")
      Select Case .Fields("Acreditacion")
        Case "A": LblTipo.Caption = .Fields("Tipo") & ", Podra retirar su dinero el: " & .Fields("Fecha_Ret")
        Case "2A": LblTipo.Caption = .Fields("Tipo") & ", Podra retirar su dinero el: " & .Fields("Fecha_Ret")
        Case "M": LblTipo.Caption = .Fields("Tipo")
      End Select
  End If
  End With
  CICliente = "CI_BLANCO"
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo = '" & CodigoCliente & "' "
  SelectAdodc AdoCuentas, sSQL
  With AdoCuentas.Recordset
   If .RecordCount > 0 Then
       LabelSocio.Caption = .Fields("Cliente")
       DCClientes.Text = LabelSocio.Caption
       CICliente = .Fields("CI_RUC")
       Cartilla_No = 0
       sSQL = "SELECT TOP 1 * " _
            & "FROM Trans_Libretas " _
            & "WHERE Cuenta_No = '" & CuentaNo.Text & "' " _
            & "ORDER BY Fecha DESC,IDT DESC,Hora DESC,ID DESC "
       SelectAdodc AdoCtaNo, sSQL
       If AdoCtaNo.Recordset.RecordCount > 0 Then
          NumeroLineas = AdoCtaNo.Recordset.Fields("ID") + 1
          'SaldoDisp = AdoCtaNo.Recordset.Fields("Saldo_Disp")
          SaldoCont = AdoCtaNo.Recordset.Fields("Saldo_Cont")
          Cartilla_No = AdoCtaNo.Recordset.Fields("Cartilla_No")
          SaldoDisp = SaldoCont - Total_Depositos - Total_Fondo_Reserva
       End If
       If Moneda_US Then OpcME.value = True Else OpcMN.value = True
       TextLinea.Text = NumeroLineas
       LabelDisponible.Caption = Format(SaldoDisp, "#,##0.00")
       LabelPorConf.Caption = Format(SaldoCont, "#,##0.00")
       LabelEnCheques.Caption = Format(Total_Depositos, "#,##0.00")
       LabelEncaje.Caption = Format(TotalEncaje, "#,##0.00")
       LabelFondoR.Caption = Format(Total_Fondo_Reserva, "#,##0.00")
       If T = Procesado Then
          If OpcCaja Then
             LabelEstado.Caption = "Apertura"
             sSQL = "SELECT Proceso " _
                  & "FROM Catalogo_Proceso " _
                  & "WHERE MidStrg(TP,1,3) = 'APE' " _
                  & "AND Nivel = 1 " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "ORDER BY Proceso "
          End If
       ElseIf T = "A" Then
          LabelEstado.Caption = "Anulada"
          MsgBox "Cuenta anulada, no podra realizar transacciones"
       Else
          LabelEstado.Caption = "Normal"
          sSQL = "SELECT Proceso " _
               & "FROM Catalogo_Proceso " _
               & "WHERE MidStrg(TP,1,3) <> 'APE' "
          If Mi_Cta Then
             sSQL = sSQL & "AND Mi_Cta <> " & Val(adFalse) & " "
          Else
             sSQL = sSQL & "AND Mi_Cta = " & Val(adFalse) & " "
          End If
          If OpcCaja Then
             sSQL = sSQL & "AND Nivel = 1 "
             'MsgBox TipoCta & vbCrLf & CFechaLong(FechaSistema) & vbCrLf & CFechaLong(Fecha_Retiro)
             If TipoDoc = "A" Or TipoDoc = "2A" Then
                If CFechaLong(FechaSistema) <= CFechaLong(Fecha_Retiro) Then
                   sSQL = sSQL & "AND DC = 'C' "
                End If
             End If
          Else
             sSQL = sSQL & "AND Nivel = 2 "
          End If
          sSQL = sSQL _
               & "AND Item = '" & NumEmpresa & "' " _
               & "ORDER BY Proceso "
       End If
       If T = "A" Then
          MBoxCuenta.SetFocus
       Else
          'MsgBox sSQL
          SelectDBList DLTP, AdoTP, sSQL, "Proceso"
          DLTP.SetFocus
       End If
   Else
       ConLibreta = False
       MsgBox "Esta cuenta no exite o esta anulada "
       LabelEstado.Caption = "NINGUNO"
       LabelSocio.Caption = "TRANSACCION SIN LIBRETA"
       ValorTotal = Val(InputBox("INGRESE EL VALOR", "SALDO ANTERIOR", 0))
       LabelDisponible.Caption = Format(ValorTotal, "#,##0.00")
       SaldoDisp = ValorTotal
       sSQL = "SELECT Proceso " _
            & "FROM Catalogo_Proceso " _
            & "WHERE MidStrg(TP,1,3) <> 'APE' " _
            & "AND Mi_Cta = 0 " _
            & "AND Nivel = 1 " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "ORDER BY Proceso "
       SelectDBList DLTP, AdoTP, sSQL, "Proceso"
       DLTP.SetFocus
   End If
  End With
  Cadena = RutaSistema & "\CI_RUC\C" & CICliente & ".GIF"
  If Dir(Cadena) <> "" Then
     PicFirma.Picture = LoadPicture(Cadena)
  Else
     Cadena = RutaSistema & "\CI_RUC\C" & CICliente & ".JPG"
     If Dir(Cadena) <> "" Then
        PicFirma.Picture = LoadPicture(Cadena)
     Else
        PicFirma.Picture = LoadPicture()
     End If
  End If
  LblCartilla.Caption = " Cartilla No. " & Format(Cartilla_No, "000000")
  LabelDispRet.Caption = Format(SaldoDisp - TotalEncaje, "#,##0.00")
End Sub

'''Private Sub Check1_Click()
''' If Check1.Value <> 0 Then DCClientes.Visible = True Else DCClientes.Visible = False
'''End Sub

Private Sub Command1_Click()
Dim Dias_Fin_Anio As Integer
  Trans_No = 21
  Monto_Total = 0
  TextoProc = DLTP.Text
  TextoValidoVar TextoProc
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE Proceso = '" & TextoProc & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCtaNo, sSQL
  With AdoCtaNo.Recordset
   If .RecordCount > 0 Then
       NivelCta = .Fields("Nivel")
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Catalogo_Apertura " _
       & "WHERE ME = 0 " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCtaNo, sSQL
  If AdoCtaNo.Recordset.RecordCount > 0 Then
     MontoAper = AdoCtaNo.Recordset.Fields("Monto_Aper")
     MontoCert = AdoCtaNo.Recordset.Fields("Monto_Cert") * NoCert
  End If
  'MsgBox MontoAper
  'If (MontoAper And MontoCert) Then T = Normal
  'MsgBox MontoAper & vbCrLf & MontoCert
  If NumCheque = "" Then NumCheque = Ninguno
  If NombreBanco = "" Then NombreBanco = Ninguno
  Select Case DC_Caja
    Case "C"
         Insertar_Montos AdoCuentas, AdoCtaNo, MBoxCuenta, 0, CCur(TxtMonto.Text), NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
         Total = CCur(TxtMonto.Text)
         Imprimir_Papeleta NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco, LabelSocio.Caption
         Imprimir_Papeleta NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco, LabelSocio.Caption, True
         If T = Procesado Then
            If CCur(TxtMonto.Text) < (MontoAper + MontoCert) Then
               MsgBox "No se debito los certificados o Gastos Bancarios revise la libreta"
            Else
                If MontoAper > 0 Then
                   NumeroLineas = NumeroLineas + 1
                   TipoProc = "N/DG"
                   Insertar_Montos AdoCuentas, AdoCtaNo, MBoxCuenta, MontoAper, 0, NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
                End If
                If MontoCert > 0 Then
                   NumeroLineas = NumeroLineas + 1
                   TipoProc = "N/DC"
                   Insertar_Montos AdoCuentas, AdoCtaNo, MBoxCuenta, MontoCert, 0, NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
                End If
            End If
         End If
         sSQL = "UPDATE Clientes_Datos_Extras " _
              & "SET T = '" & Normal & "' " _
              & "WHERE Cuenta_No = '" & MBoxCuenta.Text & "' "
         ConectarAdoExecute sSQL
    Case "D"
         Insertar_Montos AdoCuentas, AdoCtaNo, MBoxCuenta, CCur(TxtMonto.Text), 0, NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
         Total = CCur(TxtMonto)
         Imprimir_Papeleta NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco, LabelSocio.Caption
         Imprimir_Papeleta NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco, LabelSocio.Caption, True
         ValorDH = CDbl(TxtMonto)
         If TipoGrupo Then
            DetalleComp = Ninguno
           'Procesar Egreso/Diario
            NumComp = ReadSetDataNum("Egresos", True, True)
           'NumComp = ReadSetDataNum("Diario", True, True)
            IniciarAsientosAdo AdoAsientos
            Fecha_Vence = FechaSistema
            
           'MsgBox Mi_Cta
            If Mi_Cta Then
               InsertarAsientos AdoAsientos, Cta_Libretas, 0, ValorDH, 0
            Else
               InsertarAsientos AdoAsientos, Cta_Suspenso, 0, ValorDH, 0
            End If
            
            DetalleComp = "Ch. No. " & NoCheque & ", Debitado a " & LabelSocio.Caption
            InsertarAsientos AdoAsientos, SinEspaciosIzq(DCBanco), 0, 0, ValorDH
            Co.T = Normal
            Co.TP = CompEgreso
            Co.Fecha = FechaSistema
            Co.Numero = NumComp
            Co.Concepto = "(" & NumEmpresa & ") Retiro en Ch. No. " & NoCheque & ", del banco " & DCBanco & ", de la Cta. Ah. No. " & MBoxCuenta
            Co.CodigoB = CodigoCli
            Co.Efectivo = 0
            Co.Monto_Total = ValorDH
            Co.T_No = Trans_No
            Co.Item = NumEmpresa
            Co.Usuario = CodigoUsuario
            GrabarComprobante Co
            RatonNormal
            ImprimirComprobantesDe False, Co
            Imprimir_Bloque_Cheques NoCheque, NoCheque, DCBanco, NoCheque, DetalleComp
            DetalleComp = Ninguno
         End If
  End Select
  Select Case TipoProc
    Case "N/DC": Insertar_Certif AdoCuentas, AdoCtaNo, MBoxCuenta, 0, MontoCert, NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
    Case "NCCA": Insertar_Certif AdoCuentas, AdoCtaNo, MBoxCuenta, 0, CCur(TxtMonto.Text), NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
    Case "RECA": Insertar_Certif AdoCuentas, AdoCtaNo, MBoxCuenta, CCur(TxtMonto.Text), 0, NumCheque, NombreBanco, CCur(LabelDisponible.Caption)
    Case "DEPP", "DEFR"
                 Dias_Fin_Anio = CFechaLong("31/12/" & Year(FechaSistema)) - CFechaLong(FechaSistema)
                 If Dias_Fin_Anio <= 90 Then Dias_Fin_Anio = 90
                 sSQL = "SELECT * " _
                      & "FROM Trans_Bloqueos "
                 SelectAdodc AdoAux, sSQL
                 With AdoAux.Recordset
                     .AddNew
                     .Fields("T") = Normal
                     .Fields("Fecha") = FechaSistema
                     .Fields("Cuenta_No") = MBoxCuenta.Text
                     .Fields("Valor") = Redondear(CCur(TxtMonto.Text), 2)
                      If TipoProc = "DEPP" Then
                        .Fields("Cheque") = TipoProc
                        .Fields("Banco") = "DEPOSITO POR PROMOCION"
                      Else
                        .Fields("Cheque") = "FR"
                        .Fields("Banco") = "FONDO RESERVA"
                      End If
                     .Fields("Dias") = Dias_Fin_Anio
                     .Fields("Item") = NumEmpresa
                     .Update
                 End With
  End Select
  Imprimir_Libreta MBoxCuenta.Text, AdoCtaNo, 1, 7, Val(TextLinea.Text), LabelDisponible.Caption
  If Factura_No <> 0 Then
     sSQL = "UPDATE Facturas " _
          & "SET T = 'C', Saldo_MN = Saldo_MN - " & Total_Factura & " " _
          & "WHERE CodigoC = '" & CodigoCliente & "' " _
          & "AND Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     TipoDoc = MBoxCuenta.Text
    'GrabarAbonos Cta_CajaG, "DEPOSITO", Ninguno, TipoFactura, Factura_No, Total_Factura, FechaSistema
     Imprimir_Dep_Alumno NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco
     Imprimir_Dep_Alumno NivelCta, TipoGrupo, FechaSistema, TiempoTexto, MBoxCuenta.Text, TipoProc, Total, NumCheque, NombreBanco
  End If
  TxtPapeleta = "00000000"
  MBoxCuenta.SetFocus
End Sub

Private Sub Command2_Click()
  Unload FCaja
End Sub

Private Sub Command3_Click()
  Imprimir_Libreta MBoxCuenta.Text, AdoCtaNo, 1, 8, CByte(TextLinea.Text)
End Sub

Private Sub Command4_Click()
  Imprimir_Certificados MBoxCuenta.Text, AdoCtaNo, 1, 8, CByte(TextLinea.Text)
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBanco_LostFocus()
  NombreBanco = SinEspaciosIzq(DCBanco.Text)
End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCClientes_LostFocus()
  CodigoCli = CodigoCliente
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCClientes & "' ")
       If Not .EOF Then CodigoCli = .Fields("Codigo")
   End If
  End With
End Sub

Private Sub DLAlumno_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLAlumno_LostFocus()
  Factura_No = 0
  CodigoCliente = Ninguno
  With AdoAlumno.Recordset
   If .RecordCount > 0 Then
       Beneficiario = DLAlumno.Text
      .MoveFirst
      .Find ("Cliente = '" & DLAlumno.Text & "' ")
       If Not .EOF Then
          DireccionCli = .Fields("Direccion")
          Total_Factura = .Fields("Saldo_MN")
          TxtMonto.Text = Total_Factura
          Factura_No = .Fields("Factura")
          CodigoCliente = .Fields("Codigo")
          TipoFactura = .Fields("TC")
       End If
   End If
  End With
  Frame2.Visible = False
  TextLinea.SetFocus
End Sub

Private Sub DLCtas_DblClick()
  MBoxCuenta.Text = SinEspaciosIzq(DLCtas.Text)
End Sub

Private Sub DLCtas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLCtas_LostFocus()
  'MBoxCuenta.Text = SinEspaciosIzq(DLCtas.Text)
End Sub

Private Sub DLTP_GotFocus()
  DCBanco.Visible = False
  TxtBanco.Visible = False
  DCClientes.Visible = False
End Sub

Private Sub DLTP_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLTP_LostFocus()
  Factura_No = 0
  Cadena = DLTP.Text
  If Cadena = "" Then Cadena = Ninguno
  sSQL = "SELECT * " _
       & "FROM Catalogo_Proceso " _
       & "WHERE Proceso = '" & Cadena & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCtaNo, sSQL
  With AdoCtaNo.Recordset
   If .RecordCount > 0 Then
       TipoProc = .Fields("TP")
       DC_Caja = .Fields("DC")
       TipoGrupo = .Fields("Cheque")
       Label8.Visible = TipoGrupo
       Label9.Visible = TipoGrupo
       Label10.Visible = TipoGrupo
       If DC_Caja = "D" Then
          DCBanco.Visible = TipoGrupo
          DCClientes.Visible = TipoGrupo
       Else
          TxtBanco.Visible = TipoGrupo
       End If
       TxtCheque.Visible = TipoGrupo
   End If
  End With
  If TipoProc = "N/CE" Then
     Listar_Alumnos
     Frame2.Visible = True
     DLAlumno.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  TipoDoc = CompDiario
  If Supervisor = False Then
     If (CNivel(1) Or CNivel(3) Or CNivel(4) Or CNivel(6)) Then
        Command1.Enabled = False
        Command3.Enabled = False
     End If
  End If
  sSQL = "SELECT Proceso " _
       & "FROM Catalogo_Proceso "
  If OpcCaja Then
     sSQL = sSQL & "WHERE Nivel = 1 "
  Else
     sSQL = sSQL & "WHERE Nivel = 2 "
  End If
  sSQL = sSQL & "ORDER BY Proceso "
  SelectDBList DLTP, AdoTP, sSQL, "Proceso"
  sSQL = "SELECT (Codigo & ' ' & Cuenta) As NomBanco " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE TC = 'BA' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDBCombo DCBanco, AdoBanco, sSQL, "NomBanco"
 'Clientes
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Grupo <> '.' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
  
  If OpcCaja Then
     FCaja.Caption = "OPERACIONES DE CAJA "
  Else
     FCaja.Caption = "OPERACIONES DE CREDITO "
  End If
  Listar_Alumnos
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCaja
  ConectarAdodc AdoTP
  ConectarAdodc AdoAux
  ConectarAdodc AdoCtaNo
  ConectarAdodc AdoBanco
  ConectarAdodc AdoCuentas
  ConectarAdodc AdoClientes
  ConectarAdodc AdoListCtas
  ConectarAdodc AdoAsientos
  ConectarAdodc AdoAlumno
End Sub

Private Sub MBoxCuenta_GotFocus()
   TotalCredito = 0
   MBoxCuenta.Text = "00000000-0"
   LabelEnCheques.Caption = "0.00"
   LabelPorConf.Caption = "0.00"
   LabelEncaje.Caption = "0.00"
   TxtMonto.Text = ""
End Sub

Private Sub MBoxCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
   If KeyCode = vbKeyF1 Then
      CuentaNo = MBoxCuenta
      sSQL = "UPDATE Trans_Certificados " _
           & "SET IP = 0 " _
           & "WHERE Cuenta_No = '" & CuentaNo & "' "
      ConectarAdoExecute sSQL
      MsgBox "Reimpresion de Certificados activado con exito"
   End If
End Sub

Private Sub MBoxCuenta_LostFocus()
  NombreBanco = ""
  NumCheque = ""
  If MBoxCuenta.Text = "00000000-0" Then
     MsgBox "Esta cuenta no esta permitida"
     MBoxCuenta.SetFocus
  Else
     Mi_Cta = False
     Listar_Cuentas MBoxCuenta
  End If
End Sub

Private Sub TextLinea_GotFocus()
  MarcarTexto TextLinea
End Sub

Private Sub TextLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyL Then
     Cartilla_No = Val(InputBox("INGRESE EL NUEVO NUMERO DE CARTILLA No.", "RENOVACION DE CARTILLA", Cartilla_No))
     CodigoL = MidStrg(InputBox("INGRESE EL DETALLE DEL NUEVO NUMERO DE CARTILLA No.", "RENOVACION DE CARTILLA", "RENOVACION DE LIBRETA"), 1, 25)
     sSQL = "INSERT INTO Trans_Cartillas " _
          & "(Fecha,Cuenta_No,Cartilla_No,Detalle,Item) VALUES " _
          & "('" & FechaSistema & "','" & MBoxCuenta & "'," & Cartilla_No & ",'" & CodigoL & "','" & NumEmpresa & "')"
     ConectarAdoExecute sSQL
  End If
End Sub

Private Sub TxtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtBanco_LostFocus()
  TextoValido TxtBanco, , True
  NombreBanco = TxtBanco.Text
End Sub

Private Sub TxtCheque_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCheque_LostFocus()
  TextoValido TxtCheque, , True
  TxtCheque.Text = Format(Val(TxtCheque), "00000000")
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
  If (DC_Caja <> "D") Then
     FechaIni = BuscarFecha(PrimerDiaMes(FechaSistema))
     FechaFin = BuscarFecha(FechaSistema)
     sSQL = "SELECT Cuenta_No,SUM(Creditos-Debitos) As Valor " _
          & "FROM Trans_Libretas " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Cuenta_No = '" & MBoxCuenta.Text & "' " _
          & "GROUP BY Cuenta_No "
     SelectAdodc AdoCtaNo, sSQL
     '& "AND Creditos > 0 "
     TotalCredito = 0
     If AdoCtaNo.Recordset.RecordCount > 0 Then TotalCredito = AdoCtaNo.Recordset.Fields("Valor")
     TotalCredito = TotalCredito + CCur(TxtMonto.Text)
     If TotalCredito >= 2000 Then
        MsgBox "Excede el Limite por: " _
             & Format(TotalCredito, "#,##0.00") & vbCrLf _
             & "Llene el formulario."
     End If
  End If
  If SaldoDisp > 0 And ConLibreta Then
     If (DC_Caja = "D") Then
       If CCur(TxtMonto.Text) > CCur(LabelDispRet.Caption) Then
           MsgBox "Excede el saldo disponible"
           TxtMonto.SetFocus
       End If
     End If
  Else
     If (DC_Caja = "D") And (CCur(TxtMonto.Text) > SaldoDisp) Then
        MsgBox "Excede el saldo disponible"
        MBoxCuenta.SetFocus
     End If
  End If
End Sub

Private Sub TxtNombres_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNombres_LostFocus()
  TextoValido TxtNombres, , True
  I = Len(TxtNombres.Text)
  If I > 0 Then
     sSQL = "SELECT (Cuenta_No  & ' => ' & Cliente) As CuentasNo " _
          & "FROM Clientes_Datos_Extras As C,Clientes As Cl " _
          & "WHERE MidStrg(Cliente,1," & I & ") = '" & UCase(TxtNombres.Text) & "' " _
          & "AND C.T <> 'A' " _
          & "AND C.Codigo = Cl.Codigo " _
          & "ORDER BY Cuenta_No "
     SelectDBList DLCtas, AdoListCtas, sSQL, "CuentasNo"
''
''  sSQL = "SELECT (Cuenta_No  & ' => ' & Nombres & ' ' & Apellidos) As CuentasNo " _
''       & "FROM Clientes_Datos_Extras " _
''       & "WHERE SUBSTRING(Nombres,1," & I & ") = '" & UCase(TxtNombres.Text) & "' " _
''       & "AND T <> 'A' " _
''       & "ORDER BY Cuenta_No "
''  SelectDBList DLCtas, AdoListCtas, sSQL, "CuentasNo"
  End If
End Sub

Private Sub TxtPapeleta_GotFocus()
  MarcarTexto TxtPapeleta
End Sub

Private Sub TxtPapeleta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPapeleta_LostFocus()
  TextoValido TxtPapeleta, True
  TxtPapeleta = Format(CLng(TxtPapeleta), "00000000")
End Sub

Private Sub TxtRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRUC_LostFocus()
  TextoValido TxtRUC, , True
  sSQL = "SELECT (Cuenta_No  & ' => ' & Cliente) As CuentasNo " _
       & "FROM Clientes_Datos_Extras As C,Clientes AS Cl " _
       & "WHERE CI_RUC = '" & TxtRUC.Text & "' " _
       & "AND C.Tipo_Dato = 'LIBRETAS' " _
       & "AND C.T <> 'A' " _
       & "AND C.Codigo = Cl.Codigo " _
       & "ORDER BY Cuenta_No "
  SelectDBList DLCtas, AdoListCtas, sSQL, "CuentasNo"
End Sub

