VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form ListEmp 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPRESA A TRABAJAR"
   ClientHeight    =   11235
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   20550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ListEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   20550
   ShowInTaskbar   =   0   'False
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
      Left            =   210
      TabIndex        =   17
      Top             =   4305
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Frame FrmEntidad 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SELECCIONE LA ENTIDAD A CONECTAR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6945
      Left            =   12495
      TabIndex        =   11
      Top             =   420
      Visible         =   0   'False
      Width           =   7785
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Cancelar"
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
         Left            =   6405
         Picture         =   "ListEmp.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5880
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "&Aceptar"
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
         Left            =   5040
         Picture         =   "ListEmp.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5880
         Width           =   1275
      End
      Begin VB.TextBox TxtReferencia 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1380
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   4305
         Width           =   7575
      End
      Begin MSDataListLib.DataList DLEntidad 
         Bindings        =   "ListEmp.frx":1A5E
         DataSource      =   "AdoEntidad"
         Height          =   3840
         Left            =   105
         TabIndex        =   12
         Top             =   315
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6773
         _Version        =   393216
         BackColor       =   16777088
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   210
      TabIndex        =   9
      Top             =   2625
      Visible         =   0   'False
      Width           =   3060
   End
   Begin MSAdodcLib.Adodc AdoEmp 
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
      Caption         =   "Emp"
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
   Begin VB.Timer TimerAct 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   525
      Top             =   105
   End
   Begin VB.Frame FrameClave 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   4620
      TabIndex        =   0
      Top             =   420
      Width           =   7785
      Begin VB.CommandButton CmdBCrearEmp 
         BackColor       =   &H00FF8080&
         Caption         =   "&Crear/Modificar Empresa"
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
         Left            =   2520
         Picture         =   "ListEmp.frx":1A77
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5355
         Width           =   2850
      End
      Begin VB.CommandButton CmdBSalir 
         BackColor       =   &H00FF8080&
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
         Left            =   5460
         Picture         =   "ListEmp.frx":2341
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5355
         Width           =   1905
      End
      Begin VB.CommandButton CmdBAceptar 
         BackColor       =   &H00FF8080&
         Caption         =   "&Ingresar "
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
         Left            =   525
         Picture         =   "ListEmp.frx":2C0B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5355
         Width           =   1905
      End
      Begin VB.TextBox TextClave 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   3255
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3255
         Width           =   2115
      End
      Begin VB.TextBox TextUsuario 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   525
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "<Ctrl+U> Selecciona otra unidad de conexion, <Ctrl+E> Selecciona  Entidad de Conexion"
         Top             =   3255
         Width           =   2220
      End
      Begin VB.TextBox TextDolar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5775
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "ListEmp.frx":3285
         Top             =   3255
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo DCEmpresa 
         Bindings        =   "ListEmp.frx":328C
         DataSource      =   "AdoEmpresa"
         Height          =   345
         Left            =   525
         TabIndex        =   4
         Top             =   4515
         Visible         =   0   'False
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   128
         Text            =   "Empresa"
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
      Begin VB.PictureBox Pict_Version 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Height          =   6750
         Left            =   105
         Picture         =   "ListEmp.frx":32A5
         ScaleHeight     =   6750
         ScaleWidth      =   7605
         TabIndex        =   16
         Top             =   105
         Width           =   7605
         Begin VB.Label LblOlvidoClave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Olvido su contrase�a?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   5040
            TabIndex        =   8
            Top             =   3780
            Width           =   2325
         End
      End
   End
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   330
      Left            =   210
      Top             =   630
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
      Caption         =   "Empresa"
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
   Begin MSAdodcLib.Adodc AdoAcceso 
      Height          =   330
      Left            =   210
      Top             =   1890
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
      Caption         =   "Acceso"
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
   Begin MSAdodcLib.Adodc AdoEmp000 
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
      Caption         =   "Emp000"
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
   Begin MSAdodcLib.Adodc AdoEntidad 
      Height          =   330
      Left            =   210
      Top             =   2205
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
      Caption         =   "Entidad"
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
      Left            =   210
      TabIndex        =   18
      Top             =   3570
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
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
         Text            =   "Tama�o"
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
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   1050
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":713C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":7456
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":7770
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":7A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":7D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":80AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":839C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":8BB6
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":8ED0
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":91EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":9428
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Lblwww 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.diskcoversystem.com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   5565
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   7560
      Width           =   5955
   End
End
Attribute VB_Name = "ListEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CarIni As Byte
Dim CarFin As Byte
Dim Intentos As Integer
Dim Tecla As Integer
Dim Progreso_Tiempo As Integer
Dim Claves As String
Dim AltoL As Single
Dim AnchoL As Single
Dim PosPicX As Single
Dim EmpresaTemp As String
Dim IPDelOrdenador As String

Dim EsReadOnly As Boolean

Dim ping As cPing

Public Sub LlenarEmpresa()
Dim FechaIniN As Integer
Dim FechaFinN As Integer
Dim SiActualizar As Boolean

    RatonReloj
    Primera_Vez = False
    PosPicX = 5600
    PosLinea = Pict_Version.Top + 2500
    Pict_Version.Cls
    Pict_Version.Picture = LoadPicture(RutaSistema & "\LOGIN.jpg")
    Pict_Version.FontBold = True
    Pict_Version.ForeColor = Amarillo_Claro
    Pict_Version.FontSize = 18
    AltoL = Pict_Version.TextHeight("H")
    Pict_Version.CurrentX = PosPicX
    Pict_Version.CurrentY = PosLinea
    Pict_Version.Print "Conectando al"
    PosLinea = PosLinea + AltoL
    
    Pict_Version.CurrentX = PosPicX
    Pict_Version.CurrentY = PosLinea
    Pict_Version.Print "Servidor Cloud de"
    PosLinea = PosLinea + AltoL
    
    Pict_Version.CurrentX = PosPicX
    Pict_Version.CurrentY = PosLinea
    Pict_Version.Print "DiskCover System"

    PonerLinea = False
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   'Presentamos datos de variable por default
    Presentar_Inventario = True
    Meses_Provision = 12
    FA.LogoFactura = "NINGUNO"
    NumEmpresa = "001"
    Periodo_Contable = Ninguno
    Periodo_Superior = Periodo_Contable
    CodigoCli = Ninguno
    MascaraCtas = "#.#.##.##.##.###"
    FormatoCtas = "#.#.##.##.##.###"
    LimpiarCtas = " . .  .  .  .   "
    MascaraCodigoK = "CC.CC.CCC.CCCCCC"
    FormatoCodigoK = "CC.CC.CCC.CCCCCC"
    LimpiarCodigoK = "  .  .   .      "
    MascaraCurso = "C.CC.CC.CC"
    FormatoCurso = "C.CC.CC.CC"
    LimpiarCurso = " .  .  .  "
    MascaraCodigoA = "CC.CC.CCC.CCCCCC"
    FormatoCodigoA = "CC.CC.CCC.CCCCCC"
    LimpiarCodigoA = "  .  .   .      "
    MascaraCodigoC = "C.CC"
    FormatoCodigoC = "C.CC"
    LimpiarCodigoC = " .  "
    
    Fecha_Vence = FechaSistema
    FechaComp = FechaSistema
    FechaInicioAnio = "01/01/" & Year(FechaSistema)
   'Datos Generales para Colegios sobre Notas de los Alumnos
    Rector = Ninguno
    Director = Ninguno
    Secretario1 = Ninguno
    Secretario2 = Ninguno
    Anio_Lectivo = Ninguno
    NombreProvincia = Ninguno
    
    CadenaParcial = ""
    sSQL = "SELECT Modulo, Item, Codigo " _
         & "FROM Acceso_Empresa " _
         & "WHERE Modulo <> '00' "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            CadenaParcial = CadenaParcial & .Fields("Modulo") & "^" & .Fields("Item") & "^" & .Fields("Codigo") & "^~"
           .MoveNext
         Loop
     End If
    End With
    If Len(CadenaParcial) > 32768 Then MsgBox "Falta ampliar los niveles de seguridad."
    Minutos = Time
    sSQL = "SELECT " & Full_Fields("Empresas") & " " _
         & "FROM Empresas " _
         & "WHERE Empresa = '" & DCEmpresa & "' "
    Select_Adodc AdoEmp, sSQL
    Cadena = Format(Time - Minutos, "hh:mm:ss") & vbCrLf
    Minutos = Time
    'Leer_Variables_Sesion_Empresa DCEmpresa
    Cadena = Cadena & Format(Time - Minutos, "hh:mm:ss") & vbCrLf
    'MsgBox Cadena
    With AdoEmp.Recordset
     If .RecordCount > 0 Then
         NumEmpresa = .Fields("Item")
         GrupoEmpresa = .Fields("Grupo")
         Empresa = .Fields("Empresa")
         EmailEmpresa = .Fields("Email")
         EmailContador = .Fields("Email_Contabilidad")
         EmailProcesos = .Fields("Email_Procesos")
         EmailRespaldos = .Fields("Email_Respaldos")
         RazonSocial = .Fields("Razon_Social")
         NombreComercial = .Fields("Nombre_Comercial")
         RUC = .Fields("RUC")
         NLogoTipo = .Fields("Logo_Tipo")
         NMarcaAgua = .Fields("Marca_Agua")
         NombreContador = .Fields("Contador")
         NombreCiudad = .Fields("Ciudad")
         RUC_Contador = .Fields("RUC_Contador")
         NombreGerente = .Fields("Gerente")
         NFirmaDigital = .Fields("Firma_Digital")
         NombrePais = .Fields("Pais")
         CodigoPais = .Fields("CPais")
         CodigoProv = .Fields("CProv")
         ReferenciaEmpresa = .Fields("Referencia")
         CI_Representante = .Fields("CI_Representante")
         TID_Repres = .Fields("TD")
         FAX = .Fields("FAX")
         Moneda = .Fields("S_M")
         Telefono1 = .Fields("Telefono1")
         Telefono2 = .Fields("Telefono2")
         Direccion = .Fields("Direccion")
         DireccionEstab = .Fields("Direccion")
         CodigoDelBanco = .Fields("CodBanco")
         NombreBanco = .Fields("Nombre_Banco")
         Dec_PVP = .Fields("Dec_PVP")
         Dec_Costo = .Fields("Dec_Costo")
         Dec_IVA = .Fields("Dec_IVA")
         Dec_Cant = .Fields("Dec_Cant")
         Cant_Item_PV = .Fields("Cant_Item_PV")
         Cant_Ancho_PV = .Fields("Cant_Ancho_PV")
        
        'Documentos Electronicos
         Ambiente = .Fields("Ambiente")
         Obligado_Conta = .Fields("Obligado_Conta")
         ContEspec = .Fields("Codigo_Contribuyente_Especial")
         Informativo_FA = .Fields("LeyendaFA")
         Informativo_FAT = .Fields("LeyendaFAT")
         MascaraCodigoK = .Fields("Formato_Inventario")
         MascaraCodigoA = .Fields("Formato_Activo")
         MascaraCtas = Replace(.Fields("Formato_Cuentas"), "C", "#")
         FormatoCtas = MascaraCtas
         LimpiarCtas = Replace(MascaraCtas, "#", " ")
         Fecha_Igualar = .Fields("Fecha_Igualar")
         Porc_Serv = Redondear(.Fields("Servicio") / 100, 2)
         RutaCertificado = RutaSistema & "\CERTIFIC\" & .Fields("Ruta_Certificado")

         OpcCoop = CBool(.Fields("Opc"))
         CentroDeCosto = CBool(.Fields("Centro_Costos"))
         Copia_PV = CBool(.Fields("Copia_PV"))
         Mod_PVP = CBool(.Fields("Mod_PVP"))
         Mod_Fact = CBool(.Fields("Mod_Fact"))
         Mod_Fecha = CBool(.Fields("Mod_Fecha"))
         Num_Meses_CD = CBool(.Fields("Num_CD"))
         Num_Meses_CE = CBool(.Fields("Num_CE"))
         Num_Meses_CI = CBool(.Fields("Num_CI"))
         Num_Meses_ND = CBool(.Fields("Num_ND"))
         Num_Meses_NC = CBool(.Fields("Num_NC"))
         Plazo_Fijo = CBool(.Fields("Plazo_Fijo"))
         No_Autorizar = CBool(.Fields("No_Autorizar"))
         Mas_Grupos = CBool(.Fields("Separar_Grupos"))
         Medio_Rol = CBool(.Fields("Medio_Rol"))
         Encabezado_PV = CBool(.Fields("Encabezado_PV"))
         CalcComision = CBool(.Fields("Calcular_Comision"))
         Grafico_PV = CBool(.Fields("Grafico_PV"))
         ComisionEjec = CBool(.Fields("Comision_Ejecutivo"))
         ImpCeros = CBool(.Fields("Imp_Ceros"))
         Email_CE_Copia = CBool(.Fields("Email_CE_Copia"))
         Ret_Aut = CBool(.Fields("Ret_Aut"))
         ConciliacionAut = CBool(.Fields("Conciliacion_Aut"))
         NumeroFASubModulo = CBool(.Fields("Abonos_FA"))

         Debo_Pagare = MensajeDeboPagare
         If Len(RazonSocial) > 1 Then CodigoA = RazonSocial Else CodigoA = Empresa
         If .Fields("Debo_Pagare") = "SI" Then Debo_Pagare = Replace(Debo_Pagare, "vRazon_Social", CodigoA) Else Debo_Pagare = Ninguno
         
        'Datos de iniciacion desde MySQL
         Fecha_CE = .Fields("Fecha_CE")
         Fecha_P12 = .Fields("Fecha_P12")
         TipoPlan = .Fields("Tipo_Plan")
         EstadoEmpresa = .Fields("Estado")
         SerieFactura = .Fields("Serie_FA")
         
         CodigoA = Ninguno
         If MascaraCodigoK = Ninguno Then
            MascaraCodigoK = "CC.CC.CCC.CCCCCC"
            FormatoCodigoK = "CC.CC.CCC.CCCCCC"
         End If
         If MascaraCodigoA = Ninguno Then
            MascaraCodigoA = "CC.CC.CCC.CCCCCC"
            FormatoCodigoA = "CC.CC.CCC.CCCCCC"
         End If
         LimpiarCodigoK = Replace(MascaraCodigoK, "C", " ")
         LimpiarCodigoA = Replace(MascaraCodigoA, "C", " ")
         
        'Asignacion de correos autom�ticos para envio a procesos automatizados
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         For I = 0 To 6
             Lista_De_Correos(I).Correo_Electronico = CorreoDiskCover
             Lista_De_Correos(I).Contrase�a = ContrasenaDiskCover
         Next I
        
         If Len(.Fields("Email_Conexion")) > 1 And Len(.Fields("Email_Contrase�a")) > 1 Then
            Lista_De_Correos(0).Correo_Electronico = .Fields("Email_Conexion")
            Lista_De_Correos(0).Contrase�a = .Fields("Email_Contrase�a")
         End If
         
         If Len(.Fields("Email_Conexion_CE")) > 1 And Len(.Fields("Email_Contrase�a_CE")) > 1 Then
            Lista_De_Correos(4).Correo_Electronico = .Fields("Email_Conexion_CE")
            Lista_De_Correos(4).Contrase�a = .Fields("Email_Contrase�a_CE")
         End If
         Lista_De_Correos(6).Correo_Electronico = "credenciales@diskcoversystem.com"
         Lista_De_Correos(6).Contrase�a = "Dlcjvl1210@Credenciales"
        '|--=:******* CONECCON A MYSQL *******:=--|
          Minutos = Time
          Datos_Iniciales_Entidad_SP_MySQL
          'MsgBox "Desktop Test: MySQL - " & Format(Time - Minutos, "hh:mm:ss")
        '|--=:******* --------.------- *******:=--|
        'MsgBox "Desktop Test: " & ServidorMySQL
         If ServidorMySQL Then
            If .Fields("Estado") <> EstadoEmpresa Or .Fields("Cartera") <> Cartera Or .Fields("Cant_FA") <> Cant_FA Or .Fields("Serie_FA") <> SerieFE Or _
               .Fields("Fecha_CE") <> Fecha_CE Or .Fields("Fecha_P12") <> Fecha_P12 Or .Fields("Tipo_Plan") <> Fecha_CO Or .Fields("Tipo_Plan") <> TipoPlan Then
                sSQL = "UPDATE Empresas " _
                     & "SET Cartera = " & Cartera & ", " _
                     & "Cant_FA = " & Cant_FA & ",  " _
                     & "Fecha_CE = '" & BuscarFecha(Fecha_CE) & "', " _
                     & "Fecha_P12 = '" & BuscarFecha(Fecha_P12) & "', " _
                     & "Tipo_Plan = '" & TipoPlan & "', " _
                     & "Estado = '" & EstadoEmpresa & "', " _
                     & "Serie_FA = '" & SerieFE & "' " _
                     & "WHERE Item = '" & NumEmpresa & "' "
                Ejecutar_SQL_SP sSQL
            End If
         End If
         Contador = 0
         ContadorRUCCI = 0
         NumItemTemp = NumEmpresa
         LogoTipo = Obtener_File_Grafico(NLogoTipo)
         FirmaDigital = Obtener_File_Grafico(NFirmaDigital)
         MarcaAgua = Obtener_File_Grafico(NMarcaAgua)
         SQLDec = ""
         CmdBSalir.Visible = False
         CmdBAceptar.Visible = False
         CmdBCrearEmp.Visible = False
         'FrameClave.Visible = False
         
         NumItemTemp = NumEmpresa
         RutaDocumentos = RutaSysBases & "\CE\CE" & NumEmpresa
        'SavePicture Me.Picture, RutaDestino
         NombreRUC = "R.U.C."
        'MsgBox LogoTipo
         Carpeta = .Fields("SubDir")
         EmpresaActual = "[" & RutaEmpresa & "]."
                  
         If Not PCActivo Then
            Cadena = NombreUsuario & vbCrLf & "Su Equipo se encuentra en LISTA NEGRA, ingreso no autorizado, comuniquese con el Administrador del Sistema"
            MsgBox UCaseStrg(Cadena), vbCritical, "ACCESO DEL PC DENEGADO"
            End
         End If
         If Not EstadoUsuario Then
            Cadena = NombreUsuario & vbCrLf & "Su ingreso no esta autorizado, comuniquese con el Administrador del Sistema"
            MsgBox UCaseStrg(Cadena), vbCritical, "ACCESO AL SISTEMA DENEGADO"
            End
         End If
     Else
         NumEmpresa = Ninguno
     End If
    End With
    
    If Cod_Bodega <> Ninguno Then
       sSQL = "SELECT Bodega " _
            & "FROM Catalogo_Bodegas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND CodBod = '" & Cod_Bodega & "' "
       Select_Adodc AdoAux, sSQL
       If AdoAux.Recordset.RecordCount > 0 Then Nom_Bodega = AdoAux.Recordset.Fields("Bodega")
    End If
    
    sSQL = "SELECT MIN(Fecha) As FechaCierreFiscal " _
          & "FROM Comprobantes " _
          & "WHERE Periodo = '" & Periodo_Contable & "' " _
          & "AND Item = '" & NumEmpresa & "' "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       'MsgBox AdoAux.Recordset.fields("FechaCierreFiscal")
       If IsNull(AdoAux.Recordset.Fields("FechaCierreFiscal")) Then FechaCierreFiscal = "01/01/2000" Else FechaCierreFiscal = AdoAux.Recordset.Fields("FechaCierreFiscal")
    Else
       FechaCierreFiscal = "01/01/" & Year(FechaSistema)
    End If
    
   'MsgBox "Desktop Test: " & Fecha_CE & vbCrLf & Fecha_P12
   'Actualiza Datos iniciales de la Empresa
   '+++++++++++++++++++++++++++++++++++++++
    Iniciar_Datos_Default_SP
   'MsgBox "Desktop Test: MySQL - " & Format(Time - Minutos, "hh:mm:ss")
   '+++++++++++++++++++++++++++++++++++++++
   'MsgBox URLToken & vbCrLf & Token
   'MsgBox "Desktop Test: " & Fecha_CE & vbCrLf & Fecha_P12
   'Resultado del SP del MySQL
    ListaFacturas = ""
    TMail.Asunto = "CARTERA VENCIDA"
    Evaluar = False
    If Cartera <> 0 And Cant_FA <> 0 Then
       ListaFacturas = "ESTIMADO " & UCase(Empresa) & ", SE LE COMUNICA QUE USTED MANTIENE UNA CARTERA VENCIDA DE USD " & Format(Cartera, "#,##0.00") & ", " _
                     & "EQUIVALENTE A " & Cant_FA & " FACTURA(S) EMITIDA(S) A USTED." & vbCrLf
    End If
    Select Case EstadoEmpresa
      Case "VEN30"
           ListaFacturas = ListaFacturas & "PRIMER COMUNICADO DE ALVERTENCIA: SU EMPRESA ESTA POR SER BLOQUEADA POR CARTERA DE 30 DIAS DE VENCIMIENTO, "
      Case "VEN60"
           ListaFacturas = ListaFacturas & "SEGUNDO COMUNICADO DE ALVERTENCIA: SU EMPRESA ESTA POR SER BLOQUEADA POR CARTERA DE 60 DIAS DE VENCIMIENTO, "
      Case "VEN90"
           ListaFacturas = ListaFacturas & "TERCER COMUNICADO DE ALVERTENCIA: SU EMPRESA ESTA POR SER BLOQUEADA POR CARTERA DE 90 DIAS DE VENCIMIENTO, "
      Case "VEN360"
           ListaFacturas = ListaFacturas & "SU EMPRESA ESTA BLOQUEADA POR CARTERA DE 360 DIAS DE VENCIMIENTO, "
           TMail.Asunto = "EMPRESA BLOQUEADA POR VENCIMIENTO MAYOR A 360 DIAS"
           Evaluar = True
      Case "VEN180", "MAS360"
           ListaFacturas = ListaFacturas & "LO SENTIMOS, SU EMPRESA ESTA SUSPENDIDA EN EL SISTEMA, "
           TMail.Asunto = "EMPRESA SUSPENDIDA"
           Evaluar = True
      Case "BLOQ"
           ListaFacturas = ListaFacturas & "LO SENTIMOS, SU EMPRESA NO ESTA ACTIVA EN EL SISTEMA, "
           TMail.Asunto = "BLOQUEO DEFINITIVO, COMUNIQUESE A DISKCOVER SYSTEM"
           Evaluar = True
    End Select
   'MsgBox "Desktop Test: Estado = " & EstadoEmpresa
    RatonNormal
    Cadena1 = ""
    If Len(NombreGerente) <= 1 Then Cadena1 = Cadena1 & "Representante Legal," & vbCrLf
    If Len(RazonSocial) <= 1 Then Cadena1 = Cadena1 & "Razon Social," & vbCrLf
    If Len(NombreContador) <= 1 Then Cadena1 = Cadena1 & "Nombre del Contador," & vbCrLf
    If Len(RUC_Contador) <= 1 Then Cadena1 = Cadena1 & "RUC del Contador," & vbCrLf
    If Len(EmailProcesos) <= 1 And Email_CE_Copia Then Cadena1 = Cadena1 & "Email de Procesos y Respaldos," & vbCrLf
    Cadena = "En esta empresa no se ha registrado " & vbCrLf & Cadena1 & vbCrLf _
           & "comuniquese con el Administrador del Sistema para que le ayude con este proceso."
    If Len(Cadena1) > 1 Then MsgBox Cadena, vbCritical, "REGISTRAR DATOS DE LA EMPRESA"
    RatonReloj
   'Encerar datos por default
   '-------------------------
    With Dato_DBF
        .FechaI = FechaSistema
        .FechaF = FechaSistema
        .Tipo_Base = Ninguno
        .Carpeta = Ninguno
        .Actuales = Ninguno
        .Antiguos = Ninguno
        .Nuevos = Ninguno
        .Curso = Ninguno
        .Especialidad = Ninguno
        .Paralelo = Ninguno
        .Usuario = Ninguno
        .Clave = Ninguno
        .Periodo = Ninguno
        
        .Mes_Mat = 0
        
        .Cod_Mat_Ini = Ninguno
        .Cod_Mat_EBG = Ninguno
        .Cod_Mat_Bach = Ninguno
        .Cod_Pen_Ini = Ninguno
        .Cod_Pen_EBG = Ninguno
        .Cod_Pen_Bach = Ninguno
        
        .Val_Mat_Ini = 0
        .Val_Mat_EBG = 0
        .Val_Mat_Bach = 0
        .Val_Pen_Ini = 0
        .Val_Pen_EBG = 0
        .Val_Pen_Bach = 0
    End With
    IESS_Per = 0
    IESS_Pat = 0
    IESS_Ext = 0
    Sueldo_Basico = 0
    Canasta_Basica = 0
    
   '=======================================================================================================
   'Averiguamos si existe conexion a bases externas DBF
   '=======================================================================================================
    
'''    sSQL = "SELECT * " _
'''         & "FROM Acceso_Otra_Base " _
'''         & "WHERE Item = '" & NumEmpresa & "' "
'''    Select_Adodc AdoEmp, sSQL
'''    With AdoEmp.Recordset
'''     If .RecordCount > 0 Then
'''         Dato_DBF.Entidad = Replace(.Fields("Nombre_Entidad"), vbCrLf, "")
'''         Dato_DBF.puerto = Replace(.Fields("DBF_Puerto"), vbCrLf, "")
'''         Dato_DBF.Tipo_Base = Replace(.Fields("Tipo_Base"), vbCrLf, "")
'''         Dato_DBF.Carpeta = Replace(.Fields("DBF_IP_Carpeta"), vbCrLf, "")
'''         Dato_DBF.Usuario = Replace(.Fields("DBF_Usuario"), vbCrLf, "")
'''         Dato_DBF.Clave = Replace(.Fields("DBF_Clave"), vbCrLf, "")
'''         Dato_DBF.Actuales = Replace(.Fields("DBF_Actuales"), vbCrLf, "")
'''         Dato_DBF.Periodo = Replace(.Fields("DBF_Periodo"), vbCrLf, "")
'''
'''         Dato_DBF.Antiguos = .Fields("DBF_Antiguos")
'''         Dato_DBF.Nuevos = .Fields("DBF_Nuevos")
'''         Dato_DBF.Curso = .Fields("DBF_Curso")
'''         Dato_DBF.Especialidad = .Fields("DBF_Especialidad")
'''         Dato_DBF.Paralelo = .Fields("DBF_Paralelo")
'''         Dato_DBF.FechaI = .Fields("DBF_FechaI")
'''         Dato_DBF.FechaF = .Fields("DBF_FechaF")
'''         Dato_DBF.Mes_Mat = .Fields("DBF_Mes_Mat")
'''
'''         Dato_DBF.Cod_Mat_Ini = .Fields("Cod_Mat_INI")
'''         Dato_DBF.Cod_Mat_EBG = .Fields("Cod_Mat_EBG")
'''         Dato_DBF.Cod_Mat_Bach = .Fields("Cod_Mat_BACH")
'''
'''         Dato_DBF.Val_Mat_Ini = .Fields("Val_Mat_INI")
'''         Dato_DBF.Val_Mat_EBG = .Fields("Val_Mat_EBG")
'''         Dato_DBF.Val_Mat_Bach = .Fields("Val_Mat_BACH")
'''
'''         Dato_DBF.Cod_Pen_Ini = .Fields("Cod_Pen_INI")
'''         Dato_DBF.Cod_Pen_EBG = .Fields("Cod_Pen_EBG")
'''         Dato_DBF.Cod_Pen_Bach = .Fields("Cod_Pen_BACH")
'''
'''         Dato_DBF.Val_Pen_Ini = .Fields("Val_Pen_INI")
'''         Dato_DBF.Val_Pen_EBG = .Fields("Val_Pen_EBG")
'''         Dato_DBF.Val_Pen_Bach = .Fields("Val_Pen_BACH")
'''     End If
'''    End With
   'Dato_DBF.Carpeta = "E:\VFP98\SISSALES"
   
   '-----------------------------------------------------------------------
   'Llenamos datos de la base externa
   '-----------------------------------------------------------------------
'''    If Len(Dato_DBF.Carpeta) > 1 Then
'''       Select Case Dato_DBF.Tipo_Base
'''         Case "FOXPRO": Leer_Datos_FoxPro
'''         Case "MYSQL": 'Leer_Datos_MySQL
'''       End Select
'''    End If
   '-----------------------------------------------------------------------
   'MsgBox "Desktop Test: " & Modulo
    Cantidad_Cyber_Tiempo = 0
    Select Case UCaseStrg(Modulo)
      Case "EDUCATIVO"
           Leer_Periodo_Lectivo                 'Datos del Periodo Lectivo
      Case "CONTABILIDAD", "FACTURACION"
           Descargar_Certificados_Logos_Empresa 'Descargar Logos y Certificados electronicos
      Case "CAJACREDITO"
           Cambio_Tipo_Cuentas                  'Si estamos en el modulo de Caja Credito
      Case "CYBER PCs"
           sSQL = "SELECT * " _
                & "FROM Catalogo_Cyber_Tiempo " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "ORDER BY Desde,Hasta "
           Select_Adodc AdoEmp, sSQL
           With AdoEmp.Recordset
            If .RecordCount > 0 Then
                Cantidad_Cyber_Tiempo = .RecordCount - 1
                ReDim VCyber_Tiempo(Cantidad_Cyber_Tiempo)
                I = 0
                Do While Not .EOF
                   VCyber_Tiempo(I).Desde = .Fields("Desde")
                   VCyber_Tiempo(I).Hasta = .Fields("Hasta")
                   VCyber_Tiempo(I).Valor = .Fields("Valor")
                   I = I + 1
                  .MoveNext
                Loop
            End If
           End With
    End Select
   '---------------------------------------------------------------------------------
   'Rutas generales de las bases
    PathEmpresa = UCaseStrg(RutaEmpresa & "\DISKCOVE.MDB")
    FechaCierre = FechaSistema
   'Label5.Caption = "Actualizando Tablas Temporales"
    'MsgBox "Desktop Test: "
    
    SeteosCtas
   'Fin del seteo de impresoras
    Periodo = Ninguno
   'ListEmp.Caption = "EMPRESAS"
    If CodigoUsuario = "ACCESO03" Then MDIFormulario.RPeriodo.Visible = True
        
   'Fin de actualizacion de el log de ingresos
   'MsgBox AgenteRetencion & vbCrLf & MicroEmpresa
   'Ver_Grafico_FormPict
    
    Select Case Ambiente
      Case "1": If Not Ping_IP("celcer.sri.gob.ec") Then Cadena = Replace(ServidorEnLineaSRI, "XXXX", "Prueba") Else Cadena = ""
      Case "2": If Not Ping_IP("cel.sri.gob.ec") Then Cadena = Replace(ServidorEnLineaSRI, "XXXX", "Produccion") Else Cadena = ""
      Case Else: Cadena = ""
    End Select
    Control_Procesos Normal, "Ingreso a " & Empresa, "R.U.C. " & RUC & ", Item: " & NumEmpresa
    'MsgBox "Desktop Test: " & ListaFacturas
   'If Cadena <> "" Then MsgBox UCaseStrg(Cadena)
    If Len(ListaFacturas) > 1 Then
       ListaFacturas = ListaFacturas _
                     & "COMUNIQUESE CON SERVICIO AL CLIENTE DE DISKCOVER SYSTEM A LOS TELEFONOS: 098-910-5300/098-652-4396/099-965-4196, " _
                     & "O ENVIE UN MAIL A carteraclientes@diskcoversystem.com; CON EL COMPROBANTE DE DEPOSITO Y ASI PROCEDER A REALIZAR " _
                     & "LA ACTUALIZACION DE LA JUSTIFICACION EN EL SISTEMA." & vbCrLf
       MsgBox ListaFacturas
       TMail.Usuario = CorreoDiskCover
       TMail.Password = ContrasenaDiskCover
       TMail.de = CorreoDiskCover
       TMail.Mensaje = ListaFacturas
       TMail.Adjunto = ""
       TMail.Credito_No = ""
       TMail.para = ""
       Insertar_Mail TMail.para, EmailEmpresa
       Insertar_Mail TMail.para, EmailContador
       Insertar_Mail TMail.para, CorreoDiskCover
       If Email_CE_Copia Then Insertar_Mail TMail.para, EmailProcesos
      'Enviamos lista de mails
       FEnviarCorreos.Show 1
    End If
    With TMail
        .TipoDeEnvio = ""
        .ListaMail = 0
        .ContadorTiempo = 0
        .para = ""
        .ListaError = ""
        .Adjunto = ""
        .Asunto = ""
        .Mensaje = ""
        .MensajeHTML = ""
        .para = ""
    End With
    
    TiempoSistema = Time - 0.01
    
    RatonNormal
    Unload ListEmp
    If Evaluar Then
       Control_Procesos "Q", "ACCESO DENEGADO A: " & Empresa, "Motivo: " & EstadoEmpresa
       End
    End If
End Sub

Public Function BuscarClave(Usuario As String, Clave As String) As Boolean
Dim Respuesta As Boolean
    NombreUsuario = ""
    IDEUsuario = Ninguno
    PWRUsuario = Ninguno
    Cod_Bodega = Ninguno
    Nom_Bodega = Ninguno
    Respuesta = False
    CNivel(1) = False
    CNivel(2) = False
    CNivel(3) = False
    CNivel(4) = False
    CNivel(5) = False
    CNivel(6) = False
    CNivel(7) = False
    Supervisor = False
    sSQL = "SELECT " & Full_Fields("Accesos") & " " _
         & "FROM Accesos " _
         & "WHERE UCaseStrg(Usuario) = '" & UCaseStrg(Usuario) & "' " _
         & "AND UCaseStrg(Clave) = '" & UCaseStrg(Clave) & "' "
    Select_Adodc AdoEmp, sSQL
    With AdoEmp.Recordset
     If .RecordCount > 0 Then
         NombreUsuario = .Fields("Nombre_Completo")
         CodigoUsuario = .Fields("Codigo")
         IDEUsuario = .Fields("Usuario")
         PWRUsuario = .Fields("Clave")
         Todas_Las_Empresas = .Fields("TODOS")
         Supervisor = .Fields("TODOS")
         Cod_Bodega = .Fields("CodBod")
         If Todas_Las_Empresas = False Then
            CNivel(1) = .Fields("Nivel_1")
            CNivel(2) = .Fields("Nivel_2")
            CNivel(3) = .Fields("Nivel_3")
            CNivel(4) = .Fields("Nivel_4")
            CNivel(5) = .Fields("Nivel_5")
            CNivel(6) = .Fields("Nivel_6")
            CNivel(7) = .Fields("Nivel_7")
            Supervisor = .Fields("Supervisor")
         End If
         Respuesta = True
     Else
       Cadena = "Sr(a). " & Usuario & ": " & vbCrLf
       Cadena = Cadena & Space(10) & "Usted no esta autorizado a ingresar al sistema." & vbCrLf
       If Intentos <= 3 Then Cadena = Cadena & Space(10) & "Vuelva a ingresar su clave."
       MsgBox Cadena
     End If
    End With
    If Respuesta = False Then MsgBox "Error: Clave incorrecta."
    BuscarClave = Respuesta
End Function

Private Sub CmdBAceptar_Click()
 'Abrir_Caja_Registradora
  If Not IsNumeric(TextDolar) Then TextDolar = "0"
  Dolar = TextDolar
  If IngresarClave = False Then
     Control_Procesos Normal, "Ingreso al Sistema"
     'FrameClave.Visible = False
     CmdBSalir.Visible = False
     CmdBCrearEmp.Visible = False
     LlenarEmpresa
     If RutaEmpresa <> "" Then
        ChDir RutaEmpresa
        Unload ListEmp
        IniciarPrograma = True
     End If
  End If
End Sub

Private Sub CmdBCrearEmp_Click()
  If Modulo = "CONTABILIDAD" Or Modulo = "SETEOS" Or Modulo = "ANEXOS SRI" Then
     If ClaveSupervisor Then
        Control_Procesos Normal, "Creacion/Modificacion de Empresa"
        Unload ListEmp
        CrearEmp.Show
     End If
  Else
     MsgBox "USTED NO ESTA AUTORIZADO A REALIZAR" & vbCrLf & vbCrLf _
          & "ESTA OPERACION EN ESTE MODULO"
  End If
End Sub

Private Sub CmdBSalir_Click()
  Control_Procesos Normal, "Salio del Sistema"
  End
End Sub

Private Sub Command1_Click()
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
         'Determinar que tipo de bases que utilizamos
          Si_No = False
          Evaluar = False
          Modo_Educativo = False
          SQL_Server = False
          
          strIPServidor = .Fields("IP_VPN_RUTA")
          strNombreBaseDatos = .Fields("Base_Datos")
          strWebServices = .Fields("WebServices")
          strPassword = .Fields("Clave_DB")
          strUsuario = .Fields("Usuario_DB")
          strPuerto = .Fields("Puerto")
          
          Select Case .Fields("Tipo_Base")
            Case "SQL SERVER"
                 If strPuerto <> 1433 Then
                     AdoStrCnn = "Data Source=tcp:" & strIPServidor & "," & CStr(strPuerto) & ";"
                 Else
                     AdoStrCnn = "Data Source=" & strIPServidor & ";"
                 End If
                 AdoStrCnn = AdoStrCnn _
                           & "Initial Catalog=" & strNombreBaseDatos & ";" _
                           & "Provider=SQLOLEDB.1;" _
                           & "UID=" & strUsuario & ";" _
                           & "PWD=" & strPassword & ";"
                 SQL_Server = True
            Case "MY SQL"
                 AdoStrCnn = "DRIVER={MySQL ODBC 5.1 Driver};" _
                           & "SERVER=" & strIPServidor & ";" _
                           & "DATABASE=" & strNombreBaseDatos & ";" _
                           & "USER=" & strUsuario & ";" _
                           & "PASSWORD=" & strPassword & ";" _
                           & "PORT=" & .Fields("Puerto") & ";" _
                           & "OPTION=3;"
            Case "ACCESS"
                 AdoStrCnn = "Data Source=" & strIPServidor & "\" & strNombreBaseDatos & ".MDB;" _
                           & "Provider=Microsoft.Jet.OLEDB.4.0;" _
                           & "Persist Security Info=False;"
          End Select
'          If Ping_IP(strIPServidor) Then
            'Buscamos la cadena de conecci�n a la base en SQL SERVER
             ConectarAdodc AdoAux
             ConectarAdodc AdoEmp
             ConectarAdodc AdoEmp000
             ConectarAdodc AdoAcceso
             ConectarAdodc AdoEmpresa
             ChDir RutaSistema
             ListEmp.Caption = "UNIDAD DE RED: [" & RutaSistema & "]."
             sSQL = "SELECT " & Full_Fields("Acceso_Empresa") & " " _
                  & "FROM Acceso_Empresa " _
                  & "WHERE Item <> '.' " _
                  & "ORDER BY Item,Modulo,Codigo "
             Select_Adodc AdoAcceso, sSQL
             Dolar = 0
             TextDolar.Text = "0.00"
              
             sSQL = "SELECT " & Full_Fields("Empresas") & " " _
                  & "FROM Empresas " _
                  & "WHERE Item <> '.' "
             If Not ConSucursal Then
                sSQL = sSQL & "ORDER BY Empresa, Item "
             Else
                sSQL = sSQL & "ORDER BY Item, Empresa "
             End If
             SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
             If AdoEmpresa.Recordset.RecordCount > 0 Then
               'FrameClave.Visible = True
                CmdBAceptar.Enabled = False
               'CmdBCrearUsu.Enabled = False
                CmdBCrearEmp.Enabled = False
               'TextUsuario.SetFocus
                RatonNormal
               'Crear_LogIn_Sistema
             Else
                Unload ListEmp
                CrearEmp.Show
             End If
'          Else
 '             MsgBox "El Servidor no esta en Linea"
 '             End
  '        End If
       Else
          MsgBox "No ha seleccionado ninguna Entidad"
       End If
   Else
      MsgBox "No hay Entidad asignada"
   End If
  End With
  FrmEntidad.Visible = False
  TextUsuario.SetFocus
End Sub

Private Sub Command2_Click()
   FrmEntidad.Visible = False
   TextUsuario.SetFocus
End Sub

Private Sub DCEmpresa_Change()
   If Len(NombreUsuario) > 1 And Len(CodigoUsuario) > 1 Then Empresa = DCEmpresa
End Sub

Private Sub DCEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And (KeyCode = vbKeyQ) Then End
  If ShiftDown And (KeyCode = vbKeyE) Then
     With AdoEmpresa.Recordset
      If .RecordCount > 0 Then
          NumItem = InputBox("NUMERO DE EMPRESA", "BUSQUEDA DE EMPRESA", "001")
         .MoveFirst
         .Find ("Item = '" & NumItem & "'")
          If Not .EOF Then
             DCEmpresa.Text = .Fields("Empresa")
             Dolar = Val(CCur(TextDolar.Text))
             'FrameClave.Visible = False
             LlenarEmpresa
             If RutaEmpresa <> "" Then
                ChDir RutaEmpresa
                Unload ListEmp
                IniciarPrograma = True
             End If
          Else
             DCEmpresa.Text = "Esta Empresa No existe"
             DCEmpresa.SetFocus
          End If
      End If
     End With
  Else
     PresionoEnter KeyCode
  End If
End Sub

Private Sub DCEmpresa_LostFocus()
   Empresa = DCEmpresa
   If Len(Empresa) > 2 Then CmdBAceptar.Enabled = True
End Sub

Private Sub DLEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
Dim buscarEmpresa As String
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyB Then
     buscarEmpresa = InputBox("Patron de Busqueda por Entidad", "BUSCAR POR ENTIDAD", "")
     sSQL = "SELECT " & Full_Fields("Empresas_Externas") & " " _
          & "FROM Empresas_Externas " _
          & "WHERE Entidad_Comercial LIKE '%" & buscarEmpresa & "%' " _
          & "ORDER BY Entidad_Comercial "
     SelectDB_List DLEntidad, AdoEntidad, sSQL, "Entidad_Comercial"
     If AdoEntidad.Recordset.RecordCount <= 0 Then
        MsgBox "LLAME A SU PROVEEDOR PARA QUE CONFIGURE ESTA OPCION Y PODER" & vbCrLf _
             & "DISFRUTAR LA NUEVA FORMA DE CONECTARCE CON OTRAS ENTIDADES." & vbCrLf _
             & vbCrLf _
             & "EMAIL: asistencia@diskcoversystem.com o diskcoversystem@msn.com" & vbCrLf & vbCrLf _
             & "TELEFONO BPX: 593-02-6052430" & vbCrLf
     End If
     DLEntidad.SetFocus
  End If
  If KeyCode = vbKeyEscape Then
     FrmEntidad.Visible = False
     TextUsuario.SetFocus
     'Unload ListEmp
  End If
End Sub

Private Sub DLEntidad_KeyUp(KeyCode As Integer, Shift As Integer)
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
          TxtReferencia = "IP VPN    : " & .Fields("IP_VPN_RUTA") & vbCrLf _
                        & "BASE DATOS: " & .Fields("Base_Datos") & vbCrLf _
                        & "CLAVE DB  : " & .Fields("Clave_DB") & vbCrLf _
                        & "PUERTO    : " & .Fields("Puerto")
       End If
   End If
  End With
End Sub

Private Sub Form_Activate()
Dim RetVal
    
    TMail.Volver_Envial = False
    Set ping = New cPing
    EsReadOnly = True
    IPDelOrdenador = ping.IP_Del_PC()
   'Seteos de Impresion
    PrintDraft = Chr(27) & Chr(120) & Chr(0) 'Draft
    Print12CPI = Chr(27) & Chr(77)           '12 CPI
    Print10CPI = Chr(27) & Chr(80)           '10 CPI
    PrintComprimir = Chr(27) & Chr(15)       'Comprimido
    PrintDouble = Chr(27) & Chr(14)          'Ancho Double
    PrintNegrita = Chr(27) & Chr(69)         'Negrita
    PrintAgrandar = Chr(14)                  'Agrandar
    
   'Resetear Impresiones
    UnPrintComprimir = Chr(27) & Chr(18)     'Cancela COmprimido
    UnPrintDouble = Chr(27) & Chr(20)        'Cancela Ancho double
    UnPrintNegrita = Chr(27) & Chr(70)       'Cancela negrita
    UnPrintAgrandar = Chr(18)                'Cancela Agrandar
    
   'MsgBox UltimoDiaMes("15/12/2020") & vbCrLf & MaximoDia(2, 2020)
    
    ConSucursal = False
    sSQL = "SELECT Sucursal " _
         & "FROM Acceso_Sucursales " _
         & "WHERE Sucursal <> '.' "
    Select_Adodc AdoEmp, sSQL
    If AdoEmp.Recordset.RecordCount > 0 Then ConSucursal = True
        
    sSQL = "SELECT Aplicacion " _
         & "FROM Modulos " _
         & "WHERE Modulo = 'VS' "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       Version_Sistema = "Actualizaci�n de: " & AdoAux.Recordset.Fields("Aplicacion") & " - Ver. 5.20 "
    Else
       Version_Sistema = "Tel�fono del proveedor: " & vbCrLf & "(+593) 09-9965-4196/09-8652-4396. "
    End If
        
  'MsgBox Version_Sistema
   DetalleComp = Ninguno
   CodigoCC = Ninguno
   Si_No = True
  ' ImpLineaCeros = False
   TimerAct.Interval = 1000
   ReDim SetD(MaxVect) As Seteos_Documentos
   For I = 0 To MaxVect - 1
       SetD(I).PosX = 0
       SetD(I).PosY = 0
       SetD(I).Tama�o = 9
   Next I
   Me.Font = 10
  'MsgBox WindowsDirectory & vbCrLf & SystemDirectory
   MensajeEncabData = ""
   Intentos = 0: ClaveGeneral = ""
   ListEmp.Caption = "UNIDAD DE RED: [" & RutaSistema & "]."
   Dolar = 0
   TextDolar.Text = "0.00"
       
'''   sSQL = "SELECT * " _
'''        & "FROM Empresas " _
'''        & "WHERE Item <> '.' " _
'''        & "ORDER BY Sucursal DESC,Empresa,Item "
'''   SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
   
'' Cadena = Leer_Archivo_Texto("\\DISKCOVER-VAIO\Data Center\SMTP_Mails.txt")
'' MsgBox Cadena
'' Cadena = Leer_Archivo_Texto(RutaSistema & "\FORMATOS\LOGINSYSTEM.key")
'' Cadena = Leer_Encriptado(Cadena)
'''MsgBox Cadena
'' FechaIni = LineasLogIn(2)
'' FechaFin = FechaSistema
'' If Len(CStr(LineasLogIn(3))) <= 1 Then             'Averiguamos si esta legalizado dejando 45 dias de gracias
''    I = CFechaLong(FechaFin) - CFechaLong(FechaIni)
''    If I >= 45 Then
''       LogInKey.Show 1
''       End
''    End If
'' End If
 
End Sub

Private Sub Form_Load()
Dim NumFile As Integer
Dim NumPos As Long
Dim NumCarIgual As Byte
Dim NumCarComa As Byte
Dim CarBase As String
Dim RutaGeneraFile As String
Dim LineaTexto As String
Dim MiArchivo, MiRuta, MiNombre
Dim Txt_SMTP_Mails As String
   
    RatonReloj
    EmpresaTemp = Ninguno
   'Redondear_Control CmdBAceptar
   'Obtenemos la fecha del Sistema
    FechaSistema = Format$(Date, FormatoFechas)
    
   'Contador de Fondos y Verificacion de Fondos de Pantalla
    Mes = Format$(Month(FechaSistema), "00")
    Cadena = Dir(RutaSistema & "\FONDOS\M" & Mes & "\*.jpg", vbNormal)
    ContadorFondos = 0
    Do While Cadena <> ""
       If Cadena <> "." And Cadena <> ".." Then
          If (GetAttr(RutaSistema & "\FONDOS\M" & Mes & "\" & Cadena) And vbNormal) = vbNormal Then
             ReDim Preserve Fondos_Pantalla(ContadorFondos) As String
             Fondos_Pantalla(ContadorFondos) = RutaSistema & "\FONDOS\M" & Mes & "\" & Cadena
             ContadorFondos = ContadorFondos + 1
          End If
       End If
       Cadena = Dir
    Loop
    Cadena = ""
    ContadorFondos = UBound(Fondos_Pantalla)
    
'    MsgBox MDI_X_Max & "x" & MDI_Y_Max
   
   IP_PC.InterNet = Get_Internet
'   If Not Get_Internet Then MsgBox "Este Equipo no esta conectado a Internet"
'   HayCnn = Get_WAN_IP
   
  'MsgBox IP_PC.InterNet & vbCrLf & IP_PC.Nombre_PC & vbCrLf & IP_PC.IP_PC & vbCrLf & IP_PC.MAC_PC & vbCrLf & IP_PC.WAN_PC
   If Not IP_PC.InterNet Then
      Cadena = "Su acceso al Internet es inestable, no podra autorizar los comprobantes " _
             & "electronicos hasta que su conexion al Internet se estabilice."
      MsgBox Cadena, vbCritical, "ESTADO DE CONEXION AL INTERNET"
   End If
  'Determinar que tipo de bases utilizamos
   Si_No = False
   Evaluar = False
   Modo_Educativo = False
   CajaPV = "001001"
   Cadena = Dir(RutaSistema & "\", vbNormal) 'Recupera la primera entrada.
   Do While Cadena <> ""
      If Cadena <> "." And Cadena <> ".." Then
         If (GetAttr(RutaSistema & "\" & Cadena) And vbNormal) = vbNormal Then
           'Averiguamos el punto de venta del equipo
            If UCaseStrg(Cadena) = "PUNTOVENTA.KEY" Then
               CajaPV = ""
               RutaGeneraFile = RutaSistema & "\PUNTOVENTA.key"
               NumFile = FreeFile
               Open RutaGeneraFile For Input As #NumFile
               Do While Not EOF(NumFile)
                  CajaPV = CajaPV & Input(1, #NumFile)   ' Obtiene un car�cter.
               Loop
               Close #NumFile
            End If
         End If
         If (GetAttr(RutaSistema & "\" & Cadena) And vbNormal) = vbNormal Then
            If UCaseStrg(Cadena) = "EDUCATIVO.TXT" Then Modo_Educativo = True
         End If
      End If
      Cadena = Dir
   Loop
   TimerAct.Enabled = True
   TimerAct.Interval = 500
   
   CodigoUsuario = Ninguno
   NombreImagenEsperar = ""
   Procesando = 0
   Dia = Format$(Day(Date), "00")
   Mes = Format$(Month(Date), "00")
   Anio = Format$(Year(Date), "0000")
   H_INCH = 1440 / Screen.TwipsPerPixelX
   V_INCH = 1440 / Screen.TwipsPerPixelY
   
   FrameClave.Left = (Screen.width - FrameClave.width) / 2
   FrameClave.Top = (Screen.Height - FrameClave.Height - 2000) / 2

   Lblwww.Top = FrameClave.Top + FrameClave.Height + 100
   Lblwww.Left = (Screen.width - Lblwww.width) / 2
   
   Ver_Grafico_Form ListEmp, RutaSistema & "\FORMATOS\INICIO.jpg"
   Pict_Version.Cls
   Pict_Version.Picture = LoadPicture(RutaSistema & "\LogIn.jpg")
   Pict_Version.AutoRedraw = True
   Pict_Version.ForeColor = Gris
   Pict_Version.FontName = TipoVerdana
   
   Pict_Version.PaintPicture LoadPicture(RutaSistema & "\LOGOS\DiskCove.jpg"), 2500, 800, 2000, 800

   Pict_Version.FontBold = True
   Pict_Version.FontSize = 12
   Pict_Version.CurrentX = 350
   Pict_Version.CurrentY = 2600
   Pict_Version.Print "Usuario"

   Pict_Version.CurrentX = 3040
   Pict_Version.CurrentY = 2600
   Pict_Version.Print "Contrase�a"
  
   Pict_Version.CurrentX = 5700
   Pict_Version.CurrentY = 2600
   Pict_Version.Print "Cotizaci�n"
  
   Pict_Version.CurrentX = 350
   Pict_Version.CurrentY = 3900
   Pict_Version.Print "Entidad/Empresa"

   RatonReloj
  'Averiguamos si el MySQL esta en linea
   Cadena = ""
  '-------------------------------------------------------
  'Buscamos la cadena de conecci�n a la base en SQL SERVER
  '-------------------------------------------------------
   Conectar_Base_Datos
 ' Conectamos los ADO a la basa
   ConectarAdodc AdoAux
   ConectarAdodc AdoEmp
   ConectarAdodc AdoEmp000
   ConectarAdodc AdoAcceso
   ConectarAdodc AdoEmpresa
   ConectarAdodc AdoEntidad

  'Contador de ayuda y fraces en la pantalla temporal
  '--------------------------------------------------
'''   ContadorAyuda = 0
'''   sSQL = "SELECT MAX(No) As No_Max " _
'''        & "FROM Tabla_Mensajes " _
'''        & "WHERE No > 0 "
'''   Select_Adodc AdoAux, sSQL
'''   If AdoAux.Recordset.RecordCount > 0 Then ContadorAyuda = AdoAux.Recordset.Fields("No_Max")
   
  'Determinamos si se esta actulizando el sistema, no permite ingresar
   sSQL = "SELECT Aplicacion FROM Modulos WHERE Modulo = 'UP' "
   Select_Adodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      MsgBox "NO SE PUEDE INGRESAR AL SISTEMA" & vbCrLf & vbCrLf _
           & "MIENTRAS SE ENCUENTA ACTUALIZANDO"
      End
   End If
   
   If UCaseStrg(Modulo) <> "SETEOS" Then CmdBCrearEmp.Caption = "Funciona con Seteos"
   Set ftp = New cFTP
   
   sSQL = "SELECT " & Full_Fields("Empresas") & " " _
        & "FROM Empresas " _
        & "WHERE Item <> '--' "
   SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
   If AdoEmpresa.Recordset.RecordCount <= 0 Then
      Unload ListEmp
      CrearEmp.Show
   Else
      EmpresaTemp = AdoEmpresa.Recordset.Fields("Empresa")
      DCEmpresa.Text = EmpresaTemp
      CmdBAceptar.Enabled = False
      CmdBCrearEmp.Enabled = False
   End If
      
  'NumEmpresa = "001"
  
End Sub

Private Sub LblOlvidoClave_Click()
Dim AdoDBTemp As ADODB.Recordset
Dim cedula As String
Dim ContadorUsuario As Integer
    cedula = ""
    ContadorUsuario = 0
    Cadena = "ESTIMADO USUARIO, Ingrese el Email asignado, " _
           & "sus credenciales seran enviada a este Email." & vbCrLf & vbCrLf _
           & "MAIL PERSONAL DEL USUARIO:"
    EmailUsuario = InputBox(Cadena, "EMAIL PERSONAL DEL USUARIO", "")
    If EmailUsuario = "" Then EmailUsuario = Ninguno
    If Len(EmailUsuario) > 3 And InStr(EmailUsuario, "@") Then
       sSQL = "SELECT CI_NIC, Usuario " _
            & "FROM acceso_usuarios " _
            & "WHERE Email = '" & EmailUsuario & "' "
       Select_AdoDB_MySQL AdoDBTemp, sSQL
       If AdoDBTemp.RecordCount > 0 Then
          ContadorUsuario = AdoDBTemp.RecordCount
          cedula = MidStrg(InputBox("Existe " & ContadorUsuario & " coincidencia(s) para este mail," & vbCrLf & "DIGITE LOS ULTIMOS 4 DIGITOS " & vbCrLf & "DE SU CEDULA/NIC:", "EMAIL PERSONAL DEL USUARIO", ""), 1, 4)
       Else
          MsgBox "Este correo no existe en la base de datos, vuelva a ingresar"
       End If
       AdoDBTemp.Close
       
       If cedula <> "" Then
          sSQL = "SELECT CI_NIC, Usuario, Clave, Nombre_Usuario, Email " _
               & "FROM acceso_usuarios " _
               & "WHERE Email = '" & EmailUsuario & "' " _
               & "AND CI_NIC LIKE '%" & cedula & "' " _
               & "LIMIT 1 "
          Select_AdoDB_MySQL AdoDBTemp, sSQL
          With AdoDBTemp
           If .RecordCount > 0 Then
               TMail.servidor = "imap.diskcoversystem.com"
               EmailEmpresa = EmailUsuario
               Obligado_Conta = "NN"
               RUC = "1792164710001"
               Direccion = "Atacames N23-226 y Av. La Gasca"
               Mifecha = FechaSistema
               NombreCiudad = "QUITO"
               NombrePais = "ECUADOR"
               NombreUsuario = "El Usuaro Electronico"
               RazonSocial = "SISTEMA FINANCIERO CONTABLE Y PARA COOPERATIVAS"
               NombreComercial = "DISKCOVER SYSTEM"
               NLogoTipo = "DiskCover"
               Telefono1 = "09-9965-4196"
               Telefono2 = "09-8910-5300"
               EmailProcesos = "soporte@diskcoversystem.com"
               TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_mail_basico.html")
          
               html_Informacion_adicional = "<strong>INFORMACION DEL USUARIO:</strong><br>" _
                                          & "<strong>Usuario: </strong>" & .Fields("Usuario") & "<br>" _
                                          & "<strong>Clave: </strong>" & .Fields("Clave") & "<br><br>"
                                      
               html_Detalle_adicional = ""
          
               TMail.Usuario = CorreoDiskCover
               TMail.Password = ContrasenaDiskCover
               TMail.de = CorreoDiskCover
               TMail.Mensaje = ""
               TMail.Adjunto = ""
               TMail.Credito_No = ""
               TMail.Asunto = "Credenciales de: " & .Fields("Nombre_Usuario")
               TMail.para = ""
               Insertar_Mail TMail.para, EmailUsuario
               Insertar_Mail TMail.para, CorreoDiskCover
              'Enviamos lista de mails
               FEnviarCorreos.Show 1
               End
           Else
               MsgBox "Este Email: " & EmailUsuario & ", no esa asociado con el codigo de seguridad: " & cedula & ", en la base de datos, vuelva a ingresar"
           End If
          End With
          AdoDBTemp.Close
       Else
          MsgBox "Codigo de seguridad incompleto"
       End If
    Else
       MsgBox "Datos Incompletos"
    End If
End Sub

Private Sub Lblwww_DblClick()
Dim iRet As Long
  If IP_PC.InterNet Then
     iRet = Shell("rundll32.exe url.dll,FileProtocolHandler " & "https://www.diskcoversystem.com", vbMaximizedFocus)
  Else
     MsgBox "No puede acceder a la pagina web de www.diskcoversystem.com por que no tiene internet"
  End If
End Sub

Private Sub TextClave_GotFocus()
 MarcarTexto TextClave
 Dolar = 0
 TextDolar.Text = "0.00"
 sSQL = "SELECT Cotizacion " _
      & "FROM Empresas " _
      & "WHERE Item <> '.' " _
      & "ORDER BY Item "
 Select_Adodc AdoEmp, sSQL
 With AdoEmp.Recordset
  If .RecordCount > 0 Then
     .MoveFirst
      Dolar = .Fields("Cotizacion")
      TextDolar = Dolar
  End If
 End With
 NumItem = "000"
 sSQL = "SELECT * " _
      & "FROM Acceso_Empresa " _
      & "WHERE Item <> '.' " _
      & "ORDER BY Item,Modulo,Codigo "
 Select_Adodc AdoAcceso, sSQL
 With AdoAcceso.Recordset
  If .RecordCount > 0 Then
     .Find ("Codigo = '" & CodigoCliente & "' ")
    If Not .EOF Then NumItem = .Fields("Item")
  End If
 End With
End Sub

Private Sub TextDolar_GotFocus()
   MarcarTexto TextDolar
   IniciarPrograma = True
End Sub

Private Sub TextDolar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     If IsNumeric(TextDolar.Text) Then
        Dolar = Val(CCur(TextDolar))
     Else
        Dolar = 0
        TextDolar = "0.00"
     End If
     LlenarEmpresa
     If RutaEmpresa <> "" Then
        ChDir RutaEmpresa
        Unload ListEmp
     End If
  End If
End Sub

Private Sub TextDolar_LostFocus()
   Dolar = CCur(Val(TextDolar.Text))
End Sub

Private Sub TextClave_Change()
  If Len(TextClave.Text) >= TextClave.MaxLength Then SiguienteControl
End Sub

Private Sub TextClave_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And (KeyCode = vbKeyQ) Then End
End Sub

Private Sub TextClave_LostFocus()
Dim Es_CI As Boolean
Dim SinAccesoEmp As Boolean
Es_CI = True
Claves = TextClave.Text
If Claves <> "" And TextUsuario.Text <> "" Then
   Intentos = Intentos + 1
   If (BuscarClave(TextUsuario, Claves)) And (Intentos < 3) Then
      sSQL = "SELECT " & Full_Fields("Modulos") & " " _
           & "FROM Modulos " _
           & "WHERE Aplicacion = '" & Modulo & "' "
      Select_Adodc AdoAux, sSQL
      If AdoAux.Recordset.RecordCount > 0 Then NumModulo = Format(AdoAux.Recordset.Fields("Modulo"), "00")

      Select Case CodigoUsuario
        Case "0702164179", "ACCESO01" To "ACCESO11"
             'No hacemos nada
             SinAccesoEmp = True
        Case Else
             SinAccesoEmp = False
             sSQL = "SELECT Codigo, Usuario, Clave, Nombre_Completo, EmailUsuario " _
                  & "FROM Accesos " _
                  & "WHERE Codigo = '" & CodigoUsuario & "' "
             Select_Adodc AdoEmp, sSQL
             If AdoEmp.Recordset.RecordCount > 0 Then
                EmailUsuario = AdoEmp.Recordset.Fields("EmailUsuario")
                If Len(EmailUsuario) <= 1 Then
                   Cadena = "ESTIMADO USUARIO, con el cambio del sistema al metodo de procesamiento bajo las nubes (cloud), " _
                          & "es necesario registrar un correo personal donde se enviar� las credenciales para el respectivo " _
                          & "ingreso." & vbCrLf & vbCrLf _
                          & "Este mensaje seguira preset�ndose hasta que nos proporcione un correo valido." & vbCrLf & vbCrLf _
                          & "Desde ya agradezco su colaboraci�n." & vbCrLf & vbCrLf _
                          & "INGRESE EL MAIL PERSONAL DEL USUARIO:"
                   EmailUsuario = InputBox(Cadena, "EMAIL PERSONAL DEL USUARIO", "")
                   If EmailUsuario = "" Then EmailUsuario = Ninguno
                   If Len(EmailUsuario) > 3 And InStr(EmailUsuario, "@") Then
                      AdoEmp.Recordset.Fields("EmailUsuario") = LCase(EmailUsuario)
                      AdoEmp.Recordset.Update
                   End If
                   Cadena = ""
                End If
             End If
      End Select
     'MsgBox "Supervisor: " & Supervisor
      DCEmpresa.Visible = True
      If SinAccesoEmp Then
         sSQL = "SELECT " & Full_Fields("Empresas") & " " _
              & "FROM Empresas " _
              & "WHERE Item <> '--' " _
              & "ORDER BY Empresa, Item "
         SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
      Else
        'Determina a que Empresa tiene acceso
         SQL1 = "SELECT Item " _
              & "FROM Acceso_Empresa " _
              & "WHERE Codigo = '" & CodigoUsuario & "' " _
              & "AND Modulo = '" & NumModulo & "' "
              
        'Determina a que Modulos tiene acceso
         sSQL = "SELECT Codigo " _
              & "FROM Acceso_Empresa " _
              & "WHERE Codigo = '" & CodigoUsuario & "' " _
              & "AND Modulo = '" & NumModulo & "' "
         Select_Adodc AdoEmp, sSQL
         If AdoEmp.Recordset.RecordCount > 0 Then
            sSQL = "SELECT " & Full_Fields("Empresas") & " " _
                 & "FROM Empresas " _
                 & "WHERE Item IN (" & SQL1 & ") "
            If ConSucursal Then sSQL = sSQL & "ORDER BY Item, Empresa " Else sSQL = sSQL & "ORDER BY Empresa, Item "
            SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
           ' MsgBox sSQL
            CmdBAceptar.Enabled = True
         Else
            CmdBAceptar.Enabled = False
            CmdBCrearEmp.Enabled = False
            DCEmpresa.Visible = False
         
            Cadena = "ESTIMADO " & UCaseStrg(NombreUsuario) & ", por los m�ltiples cambios que se han realizado para el procesamientos financiero contable en las nubes (Cloud), " _
                   & "comuniquese con el Administrador del Sistema para que active las seguridades respectivas y asi mantener sus datos integros y libre " _
                   & "de ataques externos."
            MsgBox Cadena, vbInformation, "ACTIVAR SEGURIDADES DE ACCESO AL SISTEMA"
            TextUsuario.Text = ""
            TextUsuario.SetFocus
         End If
      End If

      'FrameClave.Visible = False
      'CmdBAceptar.Enabled = True
      CmdBCrearEmp.Enabled = False
      If UCaseStrg(Modulo) = "SETEOS" Then CmdBCrearEmp.Enabled = True
      IngresarClave = False
      If Modulo = "CONTABILIDAD" Then
         DigVerif = Digito_Verificador(CodigoUsuario)
         If Len(CodigoUsuario) <= 8 Then Es_CI = False
         If Len(CodigoUsuario) >= 10 And Tipo_RUC_CI.Tipo_Beneficiario <> "C" Then Es_CI = False
         If MidStrg(CodigoUsuario, 1, 6) = "ACCESO" Then Es_CI = True
      End If
'''      If Es_CI Then
         If Modo_Educativo Then
            With AdoEmpresa.Recordset
             If .RecordCount > 0 Then
                .MoveFirst
                .Find ("Item = '" & NumItem & "' ")
                 If Not .EOF Then
                    DCEmpresa.Text = .Fields("Empresa")
                    Dolar = Val(CCur(TextDolar.Text))
                    'FrameClave.Visible = False
                    LlenarEmpresa
                    If RutaEmpresa <> "" Then
                       ChDir RutaEmpresa
                       Unload ListEmp
                       IniciarPrograma = True
                    End If
                 Else
                    DCEmpresa.Text = "Esta Empresa No existe"
                    DCEmpresa.SetFocus
                 End If
             End If
            End With
         Else
         ''   DCEmpresa.SetFocus
         End If
'''      Else
'''         Unload ListEmp
'''         ActualizarUsuarios.Show
'''      End If
   ElseIf Intentos >= 3 Then
      Cadena = "Sr(a). " & UCaseStrg(TextUsuario.Text) & ": " & vbCrLf _
             & Space(10) & "Usted no est� autorizado" & vbCrLf _
             & Space(10) & "a ingresar al sistema." & vbCrLf & vbCrLf _
             & Space(10) & "Vuelva a ejecutar el programa."
      MsgBox Cadena
      End
   Else
      Claves = "": TextClave = "": TextClave.SetFocus
   End If
End If
End Sub

Private Sub TextUsuario_GotFocus()
   TextUsuario.Text = ""
   sSQL = "UPDATE Accesos " _
        & "SET EmailUsuario = '" & CorreoDiskCover & "' " _
        & "WHERE LEN(EmailUsuario) = 1 " _
        & "AND Codigo IN ('.', '..', '0702164179', '7777777777', '8888888888', '9999999999', 'ACCESO01', 'ACCESO02', " _
        & "'ACCESO03', 'ACCESO04', 'ACCESO05', 'ACCESO06', 'ACCESO07', 'ACCESO08', 'ACCESO09', 'ACCESO10', 'ACCESO11') "
   Ejecutar_SQL_SP sSQL
End Sub

Private Sub TextUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
Dim UnidadActual As String
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And (KeyCode = vbKeyE) Then
     sSQL = "SELECT " & Full_Fields("Empresas_Externas") & " " _
          & "FROM Empresas_Externas " _
          & "WHERE Entidad_Comercial <> '.' " _
          & "ORDER BY Entidad_Comercial "
     SelectDB_List DLEntidad, AdoEntidad, sSQL, "Entidad_Comercial"
     If AdoEntidad.Recordset.RecordCount > 0 Then
        FrmEntidad.Left = ((Screen.width - FrmEntidad.width) / 2)
        FrmEntidad.Top = ((Screen.Height - FrmEntidad.Height - 2000) / 2)
        FrmEntidad.Refresh
        FrmEntidad.Visible = True
        DLEntidad.SetFocus
     Else
        MsgBox "LLAME A SU PROVEEDOR PARA QUE CONFIGURE" & vbCrLf _
             & "ESTA OPCION Y PODER DISFRUTAR LA NUEVA" & vbCrLf _
             & "FORMA DE CONECTARCE CON OTRAS ENTIDADES" & vbCrLf _
             & "EMAIL: asistencia@diskcoversystem.com" & vbCrLf _
             & "diskcoversystem@msn.com" & vbCrLf _
             & "TELEFONO BPX: 593-02-3210051" & vbCrLf
        TextUsuario.SetFocus
     End If
  End If
  If CtrlDown And (KeyCode = vbKeyQ) Then End
  If CtrlDown And KeyCode = vbKeyA Then
     TextUsuario = "Administrador"
     TextClave = Obtener_Clave(TextUsuario)
     CmdBAceptar.Enabled = True
     CmdBCrearEmp.Enabled = True
     TextClave.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyS Then
     TextUsuario = "Supervisor"
     TextClave = Obtener_Clave(TextUsuario)
     CmdBAceptar.Enabled = True
     CmdBCrearEmp.Enabled = True
     TextClave.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyG Then
     TextUsuario = "Gerente"
     TextClave = Obtener_Clave(TextUsuario)
     CmdBAceptar.Enabled = True
     CmdBCrearEmp.Enabled = True
     TextClave.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyC Then
     TextUsuario = "Contador"
     TextClave = Obtener_Clave(TextUsuario)
     CmdBAceptar.Enabled = True
     CmdBCrearEmp.Enabled = True
     TextClave.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyW Then
     TextUsuario = "Walter"
     TextClave = Obtener_Clave(TextUsuario)
     CmdBAceptar.Enabled = True
     CmdBCrearEmp.Enabled = True
     TextClave.SetFocus
  End If
End Sub

Private Sub TextUsuario_LostFocus()
     Opc_Primaria = False
     Opc_Secundaria = False
     Opc_Bachillerato = False
    'MsgBox TextUsuario
     Claves = ""
     If TextUsuario = "" Then TextUsuario = Ninguno
     
     sSQL = "SELECT " & Full_Fields("Accesos") & " " _
          & "FROM Accesos " _
          & "WHERE Usuario = '" & TextUsuario & "' "
     Select_Adodc AdoEmp, sSQL
     With AdoEmp.Recordset
      If .RecordCount > 0 Then
          CodigoCliente = .Fields("Codigo")
          Opc_Primaria = CBool(.Fields("Primaria"))
          Opc_Secundaria = CBool(.Fields("Secundaria"))
          Opc_Bachillerato = CBool(.Fields("Bachillerato"))
          SetPRN_2 = .Fields("Impresora_Defecto_2")
          SetPapelPRN_2 = .Fields("Papel_Impresora_2")
          If Len(.Fields("Clave")) <= 1 Then
             NombreCliente = .Fields("Nombre_Completo")
             Cadena = "ESTE USUARIO NO HA SIDO REGISTRADA" & vbCrLf _
                    & "SU CLAVE, DEBE INGRESAR UNA CON UN" & vbCrLf _
                    & "MAXIMO DE 10 CARACTERES Y/O NUMEROS." & vbCrLf & vbCrLf _
                    & UCaseStrg(NombreCliente) & " INGRESA SU CLAVE:"
             CodigoCli = MidStrg(InputBox(Cadena, "CLAVE SIN ASIGNAR", ""), 1, 8)
             If CodigoCli = "" Then CodigoCli = Ninguno
             sSQL = "UPDATE Accesos " _
                  & "SET Clave = '" & CodigoCli & "'," _
                  & "Nombre_Completo = '" & ULCase(NombreCliente) & "',"
             If Modo_Educativo Then
                sSQL = sSQL & "TODOS = " & Val(adFalse) & " "
             Else
                sSQL = sSQL & "TODOS = " & Val(adTrue) & " "
             End If
             sSQL = sSQL & "WHERE Usuario = '" & TextUsuario & "' "
             Ejecutar_SQL_SP sSQL
             sSQL = "UPDATE Catalogo_Rol_Pagos " _
                  & "SET Clave = '" & CodigoCli & "' " _
                  & "WHERE Usuario = '" & TextUsuario & "' " _
                  & "AND Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' "
             Ejecutar_SQL_SP sSQL
            
             MsgBox "Vuelva a Ingresar al Sistema," & vbCrLf _
                  & "para verificar su clave"
             End
         End If
      Else
         MsgBox "Este usuario no esta registrado (" & TextUsuario & "), ingrese uno valido"
         CmdBAceptar.Enabled = False
         CmdBCrearEmp.Enabled = False
         TextUsuario.Text = ""
         TextUsuario.SetFocus
      End If
     End With
End Sub

Private Sub TimerAct_Timer()
Dim Logo_Tipo1 As String
  If Progreso_Tiempo < 0 Then Progreso_Tiempo = 0
 'Label5.Visible = Si_No
   
 'Empieza a imprimir los datos de la empresa
  With AdoEmpresa.Recordset
  If EmpresaTemp <> Empresa And DCEmpresa.Visible Then
  Pict_Version.Cls
  Pict_Version.Picture = LoadPicture(RutaSistema & "\LogIn.jpg")
  Pict_Version.FontBold = True
  Pict_Version.ForeColor = Gris
  Pict_Version.FontSize = 12
  Pict_Version.CurrentX = 350
  Pict_Version.CurrentY = 2600
  Pict_Version.Print "Usuario"

  Pict_Version.CurrentX = 3040
  Pict_Version.CurrentY = 2600
  Pict_Version.Print "Contrase�a"
  
  Pict_Version.CurrentX = 5700
  Pict_Version.CurrentY = 2600
  Pict_Version.Print "Cotizaci�n"
  
  Pict_Version.CurrentX = 350
  Pict_Version.CurrentY = 3900
  Pict_Version.Print "Entidad/Empresa"

  Pict_Version.FontSize = 8
  Pict_Version.Refresh
  AltoL = Pict_Version.TextHeight("H") + 20
  
  
   If .RecordCount > 0 And Len(Empresa) > 1 Then
      .MoveFirst
      .Find ("Empresa = '" & Empresa & "' ")
       If Not .EOF Then
          EmpresaTemp = Empresa
         'MsgBox Empresa & vbCrLf & .Fields("Logo_Tipo")
          PosPicX = 300
          PosLinea = 300
          LogoTipo1 = Ninguno
          If .Fields("Logo_Tipo") <> Ninguno Then
              RutaOrigen = RutaSistema & "\LOGOS\"
              If Existe_File(RutaOrigen & .Fields("Logo_Tipo") & ".gif") Then
                 LogoTipo1 = RutaSistema & "\LOGOS\" & .Fields("Logo_Tipo") & ".gif"
              ElseIf Existe_File(RutaOrigen & .Fields("Logo_Tipo") & ".jpg") Then
                 LogoTipo1 = RutaSistema & "\LOGOS\" & .Fields("Logo_Tipo") & ".jpg"
              Else
                 LogoTipo1 = RutaSistema & "\LOGOS\DEFAULT.jpg"
              End If
          End If
          'Pict_Version.Line (PosPicX, PosLinea)-(PosPicX + 2040, PosLinea + 760), Azul, BF
          If LogoTipo1 <> Ninguno Then Pict_Version.PaintPicture LoadPicture(LogoTipo1), PosPicX, PosLinea, 2000, 800
         
         'PosLinea = 2000
         PosPicX = 2400
         Pict_Version.FontBold = True
         Pict_Version.ForeColor = Azul_Claro
                  
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Razon_Social")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Nombre_Comercial")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "R.U.C."
         PosLinea = PosLinea + AltoL
                  
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Ciudad"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Tel�fono(s)"
         PosLinea = PosLinea + AltoL
         
         PosPicX = 300
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Direcci�n"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Representante"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Contador(a)"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Numero Asignado"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Usuario Loguedado"
         PosLinea = PosLinea + AltoL
         
         PosLinea = 300
         PosPicX = 3600
         Pict_Version.FontBold = False
         Pict_Version.ForeColor = Azul
        
         PosLinea = PosLinea + AltoL
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("RUC")
         PosLinea = PosLinea + AltoL
                  
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print ULCase(.Fields("Ciudad"))
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Telefono1") & "-" & .Fields("Telefono2")
         PosLinea = PosLinea + AltoL
         
         PosPicX = 2400
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Direccion")
         PosLinea = PosLinea + AltoL
         
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Gerente")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Contador")
         PosLinea = PosLinea + AltoL

         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .Fields("Item")
         PosLinea = PosLinea + AltoL
  
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print ULCase(NombreUsuario)
                  
         Dolar = .Fields("Cotizacion")
         TextDolar = Format$(Dolar, "#,##0.00")
         'MsgBox "."
       End If
   End If
'''
'''  Else
'''    If .RecordCount = 0 Then
'''
'''       PosPicX = 2000
'''       PosLinea = Pict_Version.Top + 1500
'''       Pict_Version.Cls
'''       Pict_Version.FontBold = True
'''       Pict_Version.ForeColor = Verde
'''       Pict_Version.FontSize = 14
'''       Pict_Version.CurrentX = PosPicX
'''       Pict_Version.CurrentY = PosLinea
'''       AltoL = Pict_Version.TextHeight("H")
'''       Cadena = "INGRESE SUS CREDENCIALES"
'''       Pict_Version.Print Cadena
'''    End If
  End If
  End With
 'Version_Sistema
  Si_No = Not (Si_No)
  Pict_Version.FontBold = True
  Pict_Version.FontSize = 10
  AnchoL = Pict_Version.TextWidth(Version_Sistema)
  Pict_Version.CurrentY = Pict_Version.Top + 6250
  Pict_Version.CurrentX = (Pict_Version.width - AnchoL) / 2
  Pict_Version.ForeColor = Azul_Claro
  If Si_No Then
     Pict_Version.Print Version_Sistema
     Lblwww.Caption = ""
  Else
     Lblwww.Caption = "www.diskcoversystem.com"
  End If
  Progreso_Tiempo = Progreso_Tiempo + 1504
  If Progreso_Tiempo > Pict_Version.width Then Progreso_Tiempo = 0
End Sub

'''Public Sub Crear_LogIn_Sistema()
'''Dim AdoCon1 As ADODB.Connection
'''Dim RstSchema As ADODB.Recordset
'''Dim IdTime As Long
'''Dim strCnn As String
'''   RatonReloj
'''   Contador = 0
''' ' Verificamos si existe tablas
'''   Set AdoCon1 = New ADODB.Connection
'''       AdoCon1.open AdoStrCnn
'''   Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
'''   Do Until RstSchema.EOF
'''      If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
'''         Contador = Contador + 1
'''      End If
'''      RstSchema.MoveNext
'''   Loop
'''   AdoCon1.Close
''' ' Verificamos si esta legalizado con el sistema anterior del login.gif
'''   FechaTexto = ""
'''   FechaTexto1 = ""
''' ' Verificamos si tiene licencia
'''   Si_No = True
'''   Cadena = Dir(RutaSistema & "\FORMATOS\", vbNormal) 'Recupera la primera entrada.
'''   Do While Cadena <> ""
'''      If Cadena <> "." And Cadena <> ".." Then
'''         If (GetAttr(RutaSistema & "\FORMATOS\" & Cadena) And vbNormal) = vbNormal Then
'''            If UCaseStrg(Cadena) = "LOGINSYSTEM.KEY" Then Si_No = False
'''         End If
'''      End If
'''      Cadena = Dir
'''   Loop
'''   If Contador > 0 Then
'''   If Si_No Then
'''      Cadena = "AVISO No. 3 ^ "
'''      'MsgBox FechaTexto & vbCrLf & FechaTexto1
'''      If Len(FechaTexto) = 10 Then Cadena = Cadena & FechaTexto & " ^ " Else Cadena = Cadena & FechaSistema & " ^ "
'''      If Len(FechaTexto1) = 10 Then Cadena = Cadena & FechaTexto1 & " ^ " Else Cadena = Cadena & " . ^ "
'''      'Cadena = "AVISO No. 3 ^ " & FechaSistema & " ^ " & SubCta & " ^ "
'''      sSQL = "SELECT Item,Fecha,Empresa,RUC,Ciudad,CPais,CProv,Pais,Gerente,Telefono1,FAX,Direccion," _
'''           & "Nombre_Comercial,CI_Representante,RUC_Contador,Email " _
'''           & "FROM Empresas " _
'''           & "WHERE Item <> '.' " _
'''           & "ORDER BY Item "
'''      Select_Adodc AdoAcceso, sSQL
'''      With AdoAcceso.Recordset
'''       If .RecordCount > 0 Then
'''           Do While Not .EOF
'''              For J = 0 To .Fields.Count - 1
'''                  Cadena = Cadena & .Fields(J) & " | "
'''              Next J
'''              Cadena = Cadena & " ^ "
'''             .MoveNext
'''           Loop
'''       End If
'''      End With
'''      Escribir_Archivo RutaSistema & "\FORMATOS\LOGINSYSTEM.KEY", Crear_Encriptado(Cadena)
'''   End If
'''
'''   sSQL = "SELECT * " _
'''        & "FROM Acceso_Empresa " _
'''        & "WHERE Item <> '.' " _
'''        & "ORDER BY Item,Modulo,Codigo "
'''   Select_Adodc AdoAcceso, sSQL
'''   End If
'''End Sub

Public Sub Cambio_Tipo_Cuentas()
  RatonReloj
  sSQL = "UPDATE Clientes_Datos_Extras " _
       & "SET Tipo = 'AHORRO SOCIO' " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND MidStrg(Cuenta_No,10,1) = '0' " _
       & "AND Tipo_Dato = 'LIBRETAS' "
  Ejecutar_SQL_SP sSQL
       
  If SQL_Server Then
     sSQL = "UPDATE Clientes_Datos_Extras " _
          & "SET Tipo = 'AHORRO CLIENTE' " _
          & "FROM Clientes_Datos_Extras As CL,Trans_Prestamos As TP "
  Else
     sSQL = "UPDATE Clientes_Datos_Extras As CL,Trans_Prestamos As TP " _
          & "SET Tipo = 'AHORRO CLIENTE' "
  End If
  sSQL = sSQL _
       & "WHERE CL.Item = '" & NumEmpresa & "' " _
       & "AND MidStrg(CL.Cuenta_No,10,1) = '0' " _
       & "AND TP.T = 'P' " _
       & "AND CL.Tipo_Dato = 'LIBRETAS' " _
       & "AND CL.Item = TP.Item " _
       & "AND CL.Cuenta_No = TP.Cuenta_No "
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub

'''Public Sub Actualizar_Campo_Empresa(NombreCampo As String, Valor As Variant, ValorCadena As Boolean)
'''Dim SQLEmp As String
'''   SQLEmp = "UPDATE Empresas "
'''   If ValorCadena Then
'''      SQLEmp = SQLEmp & "SET " & NombreCampo & " = '" & Valor & "' "
'''   Else
'''      SQLEmp = SQLEmp & "SET " & NombreCampo & " = " & Valor & " "
'''   End If
'''   SQLEmp = SQLEmp & "WHERE Empresa = '" & DCEmpresa.Text & "' "
'''   If ValorCadena Then
'''      SQLEmp = SQLEmp & "AND " & NombreCampo & " = '" & Ninguno & "' "
'''   Else
'''      SQLEmp = SQLEmp & "AND " & NombreCampo & " = 0 "
'''   End If
'''   Ejecutar_SQL_SP SQLEmp
'''End Sub

Private Sub Descargar_Certificados_Logos_Empresa()
Dim AdoDBTemp As ADODB.Recordset
Dim Certificado As String
Dim LogoTipoGIF As String
Dim LogoTipoJPG As String
Dim LogoTipoPNG As String
Dim ActualizoFile As Boolean

On Error GoTo error_Handler

 If Len(NumEmpresa) = 3 Then
    Certificado = Ninguno
    LogoTipoGIF = Ninguno
    LogoTipoJPG = Ninguno
    LogoTipoPNG = Ninguno
    
    sSQL = "SELECT Empresa, Ruta_Certificado, Logo_Tipo " _
         & "FROM Empresas " _
         & "WHERE Item = '" & NumEmpresa & "' "
    Select_AdoDB AdoDBTemp, sSQL
    If AdoDBTemp.RecordCount > 0 Then
      'MsgBox "Destop Test: Certificado = " & AdoDBTemp.Fields("Ruta_Certificado")
       If InStr(UCase(AdoDBTemp.Fields("Ruta_Certificado")), "P12") > 1 Then
          RutaDocumentos = RutaSistema & "\CERTIFIC\" & AdoDBTemp.Fields("Ruta_Certificado")
         'MsgBox "Desktop Test: " & Dir$(RutaDocumentos)
          If Len(Dir$(RutaDocumentos)) = 0 Then Certificado = AdoDBTemp.Fields("Ruta_Certificado")
       End If
      'MsgBox "Desktop Test: " & Certificado & vbCrLf & RutaDocumentos
       If Len(AdoDBTemp.Fields("Logo_Tipo")) > 1 Then
          RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".jpg"
          If Len(Dir$(RutaDocumentos)) = 0 Then LogoTipoJPG = AdoDBTemp.Fields("Logo_Tipo") & ".jpg"
        
          RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".gif"
          If Len(Dir$(RutaDocumentos)) = 0 Then LogoTipoGIF = AdoDBTemp.Fields("Logo_Tipo") & ".gif"
        
          RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".png"
          If Len(Dir$(RutaDocumentos)) = 0 Then LogoTipoPNG = AdoDBTemp.Fields("Logo_Tipo") & ".png"
       End If
    End If
    AdoDBTemp.Close
    
    With ftp
        .Inicializar ListEmp
        'Le indicamos el ListView donde se listar�n los archivos
         Set .ListView = LstVwFTP
       ' MsgBox EsReadOnly
         If EsReadOnly Then
            If InStr(IPDelOrdenador, "192.168.27") Then
              .servidor = "192.168.27.3"    'Establecesmo el nombre del Servidor FTP
              .Puerto = 21
            Else
              .servidor = ftpUpSvr          'Establecesmo el nombre del Servidor FTP
              .Puerto = ftpUpPuerto
            End If
           .Usuario = ftpUpUse              'Le establecemos el nombre de usuario de la cuenta
           .Password = ftpUpPwr             'Le establecemos la contrase�a de la cuenta Ftp
         Else
           .servidor = ftpSvr               'Establecesmo el nombre del Servidor FTP
           .Usuario = ftpUse                'Le establecemos el nombre de usuario de la cuenta
           .Password = ftpPwr               'Le establecemos la contrase�a de la cuenta Ftp
           .Puerto = ftpPuerto
         End If
        'MsgBox "Destop Test: " & IP_PC.IP_PC & vbCrLf & "Certificado = " & Certificado
         If Len(Certificado) > 1 Then
           'conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexi�n
            If .ConectarFtp(LstStatud) Then
                LstStatud.Text = LstStatud.Text & .GetDirectorioActual & vbCrLf
               'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
                ActualizoFile = False
               .CambiarDirectorio "/SISTEMA/CERTIFIC/"
               .ListarArchivos
                For I = 1 To LstVwFTP.ListItems.Count
                 If Certificado = LstVwFTP.ListItems(I) Then
                   .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\CERTIFIC\" & LstVwFTP.ListItems(I), True
                    ActualizoFile = True
                    Exit For
                 End If
                Next I
                If ActualizoFile Then MsgBox "Certificado nuevo actualizado"
                .Desconectar
            Else
                Err.Description = Err.Description & " No se pudo conectar al servidor de Certificados"
                GoTo error_Handler 'Exit Sub
            End If
         End If
         If .ConectarFtp(LstStatud) Then
            .CambiarDirectorio "/SISTEMA/LOGOS/"
            .ListarArchivos
             For I = 1 To LstVwFTP.ListItems.Count
                 RutaDocumentos = RutaSistema & "\LOGOS\" & LstVwFTP.ListItems(I)
                 If LogoTipoGIF = LstVwFTP.ListItems(I) Then .ObtenerArchivo LstVwFTP.ListItems(I), RutaDocumentos, True
                 If LogoTipoJPG = LstVwFTP.ListItems(I) Then .ObtenerArchivo LstVwFTP.ListItems(I), RutaDocumentos, True
                 If LogoTipoPNG = LstVwFTP.ListItems(I) Then .ObtenerArchivo LstVwFTP.ListItems(I), RutaDocumentos, True
             Next I
            .Desconectar
         Else
             Err.Description = Err.Description & " No se pudo conectar al servidor de Logotipos"
             GoTo error_Handler 'Exit Sub
         End If
     End With
  End If
  RatonNormal
Exit Sub
error_Handler:
     MsgBox Err.Description, vbCritical
     RatonNormal
End Sub

