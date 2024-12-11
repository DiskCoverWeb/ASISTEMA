VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form ListEmp 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPRESA A TRABAJAR"
   ClientHeight    =   11235
   ClientLeft      =   30
   ClientTop       =   330
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
      Left            =   105
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.Frame FrmEntidad 
      BackColor       =   &H0000C0C0&
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
      Height          =   5370
      Left            =   7350
      TabIndex        =   5
      Top             =   5250
      Visible         =   0   'False
      Width           =   8625
      Begin VB.CommandButton Command2 
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
         Left            =   7245
         Picture         =   "ListEmp.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4410
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
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
         Left            =   5880
         Picture         =   "ListEmp.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4410
         Width           =   1275
      End
      Begin VB.TextBox TxtReferencia 
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   9
         Top             =   2940
         Width           =   8415
      End
      Begin MSDataListLib.DataList DLEntidad 
         Bindings        =   "ListEmp.frx":1A5E
         DataSource      =   "AdoEntidad"
         Height          =   2535
         Left            =   105
         TabIndex        =   6
         Top             =   315
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4471
         _Version        =   393216
         BackColor       =   16777152
         ForeColor       =   12582912
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
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   210
      TabIndex        =   3
      Top             =   2625
      Visible         =   0   'False
      Width           =   2010
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
      Height          =   5370
      Left            =   3150
      TabIndex        =   0
      Top             =   420
      Width           =   13455
      Begin VB.CommandButton CmdBCrearEmp 
         BackColor       =   &H00C0E0FF&
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
         Left            =   10080
         Picture         =   "ListEmp.frx":1A77
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3990
         Width           =   2535
      End
      Begin VB.CommandButton CmdBSalir 
         BackColor       =   &H00C0E0FF&
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
         Left            =   5880
         Picture         =   "ListEmp.frx":2341
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3990
         Width           =   2535
      End
      Begin VB.CommandButton CmdBAceptar 
         BackColor       =   &H00C0E0FF&
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
         Height          =   750
         Left            =   1155
         Picture         =   "ListEmp.frx":2C0B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3780
         Width           =   3375
      End
      Begin VB.TextBox TextClave 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   945
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2415
         Width           =   3690
      End
      Begin VB.TextBox TextUsuario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   945
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "<Ctrl+U> Selecciona otra unidad de conexion, <Ctrl+E> Selecciona  Entidad de Conexion"
         Top             =   1365
         Width           =   3690
      End
      Begin VB.TextBox TextDolar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2730
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "ListEmp.frx":3285
         Top             =   3150
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo DCEmpresa 
         Bindings        =   "ListEmp.frx":328C
         DataSource      =   "AdoEmpresa"
         Height          =   345
         Left            =   5985
         TabIndex        =   11
         Top             =   1470
         Visible         =   0   'False
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   609
         _Version        =   393216
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
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         Height          =   5145
         Left            =   105
         Picture         =   "ListEmp.frx":32A5
         ScaleHeight     =   5145
         ScaleWidth      =   13200
         TabIndex        =   12
         Top             =   105
         Width           =   13200
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
      Left            =   105
      TabIndex        =   17
      Top             =   4410
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
         Text            =   "Tamaño"
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
            Picture         =   "ListEmp.frx":A9EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":AD08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":B022
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":B328
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":B642
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":B95C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":BC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":C468
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":C782
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":CA9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListEmp.frx":CCDA
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
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   4725
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   7560
      Width           =   8580
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
            CadenaParcial = CadenaParcial & .fields("Modulo") & "^" & .fields("Item") & "^" & .fields("Codigo") & "^~"
           .MoveNext
         Loop
     End If
    End With
    If Len(CadenaParcial) > 32768 Then MsgBox "Falta ampliar los niveles de seguridad."

    sSQL = "SELECT " & Full_Fields("Empresas") & " " _
         & "FROM Empresas " _
         & "WHERE Empresa = '" & DCEmpresa & "' "
    Select_Adodc AdoEmp, sSQL
    With AdoEmp.Recordset
     If .RecordCount > 0 Then
         NumEmpresa = .fields("Item")
         GrupoEmpresa = .fields("Grupo")
         Empresa = .fields("Empresa")
         EstadoEmpresa = .fields("Estado")
         Fecha_CE = .fields("Fecha_CE")
         Fecha_P12 = .fields("Fecha_P12")
         EmailEmpresa = .fields("Email")
         EmailContador = .fields("Email_Contabilidad")
         EmailProcesos = .fields("Email_Procesos")
         EmailRespaldos = .fields("Email_Respaldos")
         RazonSocial = .fields("Razon_Social")
         NombreComercial = .fields("Nombre_Comercial")
         RUC = .fields("RUC")
         NLogoTipo = .fields("Logo_Tipo")
         NMarcaAgua = .fields("Marca_Agua")
         NombreContador = .fields("Contador")
         NombreCiudad = .fields("Ciudad")
         RUC_Contador = .fields("RUC_Contador")
         NombreGerente = .fields("Gerente")
         
         TipoPlan = .fields("Tipo_Plan")
         NFirmaDigital = .fields("Firma_Digital")
         NombrePais = .fields("Pais")
         CodigoPais = .fields("CPais")
         CodigoProv = .fields("CProv")
         ReferenciaEmpresa = .fields("Referencia")
         CI_Representante = .fields("CI_Representante")
         TID_Repres = .fields("TD")
         FAX = .fields("FAX")
         Moneda = .fields("S_M")
         Telefono1 = .fields("Telefono1")
         Telefono2 = .fields("Telefono2")
         Direccion = .fields("Direccion")
         DireccionEstab = .fields("Direccion")
         CodigoDelBanco = .fields("CodBanco")
         NombreBanco = .fields("Nombre_Banco")
         Dec_PVP = .fields("Dec_PVP")
         Dec_Costo = .fields("Dec_Costo")
         Dec_IVA = .fields("Dec_IVA")
         Dec_Cant = .fields("Dec_Cant")
         Cant_Item_PV = .fields("Cant_Item_PV")
         Cant_Ancho_PV = .fields("Cant_Ancho_PV")
        
        'Documentos Electronicos
         Ambiente = .fields("Ambiente")
         Obligado_Conta = .fields("Obligado_Conta")
         ContEspec = .fields("Codigo_Contribuyente_Especial")
         Informativo_FA = .fields("LeyendaFA")
         Informativo_FAT = .fields("LeyendaFAT")
         MascaraCodigoK = .fields("Formato_Inventario")
         MascaraCodigoA = .fields("Formato_Activo")
         MascaraCtas = Replace(.fields("Formato_Cuentas"), "C", "#")
         FormatoCtas = MascaraCtas
         LimpiarCtas = Replace(MascaraCtas, "#", " ")
         Fecha_Igualar = .fields("Fecha_Igualar")
         Porc_Serv = Redondear(.fields("Servicio") / 100, 2)
         SerieFactura = .fields("Serie_FA")
         RutaCertificado = .fields("Ruta_Certificado")
         
         OpcCoop = CBool(.fields("Opc"))
         CentroDeCosto = CBool(.fields("Centro_Costos"))
         Copia_PV = CBool(.fields("Copia_PV"))
         Mod_Fact = CBool(.fields("Mod_Fact"))
         Mod_Fecha = CBool(.fields("Mod_Fecha"))
         Num_Meses_CD = CBool(.fields("Num_CD"))
         Num_Meses_CE = CBool(.fields("Num_CE"))
         Num_Meses_CI = CBool(.fields("Num_CI"))
         Num_Meses_ND = CBool(.fields("Num_ND"))
         Num_Meses_NC = CBool(.fields("Num_NC"))
         Plazo_Fijo = CBool(.fields("Plazo_Fijo"))
         No_Autorizar = CBool(.fields("No_Autorizar"))
         Mas_Grupos = CBool(.fields("Separar_Grupos"))
         Medio_Rol = CBool(.fields("Medio_Rol"))
         Encabezado_PV = CBool(.fields("Encabezado_PV"))
         CalcComision = CBool(.fields("Calcular_Comision"))
         Grafico_PV = CBool(.fields("Grafico_PV"))
         ComisionEjec = CBool(.fields("Comision_Ejecutivo"))
         ImpCeros = CBool(.fields("Imp_Ceros"))
         Email_CE_Copia = CBool(.fields("Email_CE_Copia"))
         Ret_Aut = CBool(.fields("Ret_Aut"))
         'ConciliacionAut = CBool(.fields("Conciliacion_Aut"))

         Debo_Pagare = MensajeDeboPagare
         If Len(RazonSocial) > 1 Then CodigoA = RazonSocial Else CodigoA = Empresa
         If .fields("Debo_Pagare") = "SI" Then Debo_Pagare = Replace(Debo_Pagare, "Razon_Social", CodigoA) Else Debo_Pagare = Ninguno
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
         
        'Asignacion de correos automáticos para envio a procesos automatizados
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         For I = 0 To 6
             Lista_De_Correos(I).Correo_Electronico = CorreoDiskCover
             Lista_De_Correos(I).Contraseña = ContrasenaDiskCover
         Next I
        
         If Len(.fields("Email_Conexion")) > 1 And Len(.fields("Email_Contraseña")) > 1 Then
            Lista_De_Correos(0).Correo_Electronico = .fields("Email_Conexion")
            Lista_De_Correos(0).Contraseña = .fields("Email_Contraseña")
         End If
         
         If Len(.fields("Email_Conexion_CE")) > 1 And Len(.fields("Email_Contraseña_CE")) > 1 Then
            Lista_De_Correos(4).Correo_Electronico = .fields("Email_Conexion_CE")
            Lista_De_Correos(4).Contraseña = .fields("Email_Contraseña_CE")
         End If
         Lista_De_Correos(6).Correo_Electronico = "credenciales@diskcoversystem.com"
         Lista_De_Correos(6).Contraseña = "Dlcjvl1210@Credenciales"
        '|--=:******* CONECCON A MYSQL *******:=--|
          Datos_Iniciales_Entidad_SP_MySQL
        '|--=:******* --------.------- *******:=--|
         If ServidorMySQL Then
            If .fields("Estado") <> EstadoEmpresa Or .fields("Cartera") <> Cartera Or .fields("Cant_FA") <> Cant_FA Or _
               .fields("Fecha_CE") <> Fecha_CE Or .fields("Fecha_P12") <> Fecha_P12 Or .fields("Tipo_Plan") <> TipoPlan Or _
               .fields("Serie_FA") <> SerieFE Then
               .fields("Cartera") = Cartera
               .fields("Cant_FA") = Cant_FA
               .fields("Fecha_CE") = Fecha_CE
               .fields("Fecha_P12") = Fecha_P12
               .fields("Tipo_Plan") = TipoPlan
               .fields("Estado") = EstadoEmpresa
               .fields("Serie_FA") = SerieFE
               .Update
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
         Carpeta = .fields("SubDir")
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
       If AdoAux.Recordset.RecordCount > 0 Then Nom_Bodega = AdoAux.Recordset.fields("Bodega")
    End If
   'MsgBox Fecha_CE & vbCrLf & Fecha_P12
   'Actualiza Datos iniciales de la Empresa
   '+++++++++++++++++++++++++++++++++++++++
    Iniciar_Datos_Default_SP
    
   '+++++++++++++++++++++++++++++++++++++++
   'MsgBox URLToken & vbCrLf & Token
    
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
    Cantidad_Cyber_Tiempo = 0
    Select Case UCaseStrg(Modulo)
      Case "EDUCATIVO"
           Leer_Periodo_Lectivo             'Datos del Periodo Lectivo
      Case "SETEOS"
           Descargar_FTP_Certificados_Logos 'Descargar Logos y Certificados electronicos
      Case "CAJACREDITO"
           Cambio_Tipo_Cuentas              'Si estamos en el modulo de Caja Credito
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
                   VCyber_Tiempo(I).Desde = .fields("Desde")
                   VCyber_Tiempo(I).Hasta = .fields("Hasta")
                   VCyber_Tiempo(I).Valor = .fields("Valor")
                   I = I + 1
                  .MoveNext
                Loop
            End If
           End With
    End Select
   '-----------------------------------------------------------------------
   'Rutas generales de las bases
    PathEmpresa = UCaseStrg(RutaEmpresa & "\DISKCOVE.MDB")
    FechaCierre = FechaSistema
   'Label5.Caption = "Actualizando Tablas Temporales"
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
         NombreUsuario = .fields("Nombre_Completo")
         CodigoUsuario = .fields("Codigo")
         IDEUsuario = .fields("Usuario")
         PWRUsuario = .fields("Clave")
         Todas_Las_Empresas = .fields("TODOS")
         Supervisor = .fields("TODOS")
         Cod_Bodega = .fields("CodBod")
         If Todas_Las_Empresas = False Then
            CNivel(1) = .fields("Nivel_1")
            CNivel(2) = .fields("Nivel_2")
            CNivel(3) = .fields("Nivel_3")
            CNivel(4) = .fields("Nivel_4")
            CNivel(5) = .fields("Nivel_5")
            CNivel(6) = .fields("Nivel_6")
            CNivel(7) = .fields("Nivel_7")
            Supervisor = .fields("Supervisor")
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
          
          strIPServidor = .fields("IP_VPN_RUTA")
          strNombreBaseDatos = .fields("Base_Datos")
          strWebServices = .fields("WebServices")
          strPassword = .fields("Clave_DB")
          strUsuario = .fields("Usuario_DB")
          strPuerto = .fields("Puerto")
          
          Select Case .fields("Tipo_Base")
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
                           & "PORT=" & .fields("Puerto") & ";" _
                           & "OPTION=3;"
            Case "ACCESS"
                 AdoStrCnn = "Data Source=" & strIPServidor & "\" & strNombreBaseDatos & ".MDB;" _
                           & "Provider=Microsoft.Jet.OLEDB.4.0;" _
                           & "Persist Security Info=False;"
          End Select
'          If Ping_IP(strIPServidor) Then
            'Buscamos la cadena de conección a la base en SQL SERVER
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
             DCEmpresa.Text = .fields("Empresa")
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
          TxtReferencia = "IP VPN    : " & .fields("IP_VPN_RUTA") & vbCrLf _
                        & "BASE DATOS: " & .fields("Base_Datos") & vbCrLf _
                        & "CLAVE DB  : " & .fields("Clave_DB") & vbCrLf _
                        & "PUERTO    : " & .fields("Puerto")
       End If
   End If
  End With
End Sub

Private Sub Form_Activate()
Dim RetVal
    
    TMail.Volver_Envial = False
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
       Version_Sistema = "Actualización de: " & AdoAux.Recordset.fields("Aplicacion") & " - Ver. 5.20 "
    Else
       Version_Sistema = "Teléfono del proveedor: " & vbCrLf & "(+593) 09-9965-4196/09-8652-4396. "
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
       SetD(I).Tamaño = 9
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
   'Obtenemos la fecha del Sistema
    FechaSistema = Format$(date, FormatoFechas)
    
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
                  CajaPV = CajaPV & Input(1, #NumFile)   ' Obtiene un carácter.
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
   Dia = Format$(Day(date), "00")
   Mes = Format$(Month(date), "00")
   Anio = Format$(Year(date), "0000")
   H_INCH = 1440 / Screen.TwipsPerPixelX
   V_INCH = 1440 / Screen.TwipsPerPixelY
   
   FrameClave.Left = (Screen.width - FrameClave.width) / 2
   FrameClave.Top = (Screen.Height - FrameClave.Height) / 2

   Lblwww.Top = FrameClave.Top + FrameClave.Height + 100
   Lblwww.Left = (Screen.width - Lblwww.width) / 2
   
   Ver_Grafico_Form ListEmp, RutaSistema & "\FORMATOS\INICIO.jpg"
   Pict_Version.Cls
   Pict_Version.Picture = LoadPicture(RutaSistema & "\LOGIN.jpg")
   Pict_Version.AutoRedraw = True
   Pict_Version.ForeColor = Amarillo_Claro
   Pict_Version.FontName = TipoArial

   RatonReloj
  'Averiguamos si el MySQL esta en linea
   Cadena = ""
  '-------------------------------------------------------
  'Buscamos la cadena de conección a la base en SQL SERVER
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
   
   sSQL = "SELECT " & Full_Fields("Empresas") & " " _
        & "FROM Empresas " _
        & "WHERE Item <> '--' "
   SelectDB_Combo DCEmpresa, AdoEmpresa, sSQL, "Empresa"
   If AdoEmpresa.Recordset.RecordCount <= 0 Then
      Unload ListEmp
      CrearEmp.Show
   Else
      CmdBAceptar.Enabled = False
      CmdBCrearEmp.Enabled = False
   End If
      
  'NumEmpresa = "001"
   If UCaseStrg(Modulo) <> "SETEOS" Then CmdBCrearEmp.Caption = "Funciona con Seteos"
   Set ftp = New cFTP
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
      Dolar = .fields("Cotizacion")
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
    If Not .EOF Then NumItem = .fields("Item")
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
      If AdoAux.Recordset.RecordCount > 0 Then NumModulo = Format(AdoAux.Recordset.fields("Modulo"), "00")

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
                EmailUsuario = AdoEmp.Recordset.fields("EmailUsuario")
                If Len(EmailUsuario) <= 1 Then
                   Cadena = "ESTIMADO USUARIO, con el cambio del sistema al metodo de procesamiento bajo las nubes (cloud), " _
                          & "es necesario registrar un correo personal donde se enviará las credenciales para el respectivo " _
                          & "ingreso." & vbCrLf & vbCrLf _
                          & "Este mensaje seguira presetándose hasta que nos proporcione un correo valido." & vbCrLf & vbCrLf _
                          & "Desde ya agradezco su colaboración." & vbCrLf & vbCrLf _
                          & "INGRESE EL MAIL PERSONAL DEL USUARIO:"
                   EmailUsuario = InputBox(Cadena, "EMAIL PERSONAL DEL USUARIO", "")
                   If EmailUsuario = "" Then EmailUsuario = Ninguno
                   If Len(EmailUsuario) > 3 And InStr(EmailUsuario, "@") Then
                      AdoEmp.Recordset.fields("EmailUsuario") = LCase(EmailUsuario)
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
            CmdBAceptar.Enabled = True
         Else
            CmdBAceptar.Enabled = False
            CmdBCrearEmp.Enabled = False
            DCEmpresa.Visible = False
         
            Cadena = "ESTIMADO " & UCaseStrg(NombreUsuario) & ", por los múltiples cambios que se han realizado para el procesamientos financiero contable en las nubes (Cloud), " _
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
                    DCEmpresa.Text = .fields("Empresa")
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
             & Space(10) & "Usted no está autorizado" & vbCrLf _
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
        FrmEntidad.Top = ((Screen.Height - FrmEntidad.Height) / 2)
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
          CodigoCliente = .fields("Codigo")
          Opc_Primaria = CBool(.fields("Primaria"))
          Opc_Secundaria = CBool(.fields("Secundaria"))
          Opc_Bachillerato = CBool(.fields("Bachillerato"))
          SetPRN_2 = .fields("Impresora_Defecto_2")
          SetPapelPRN_2 = .fields("Papel_Impresora_2")
          If Len(.fields("Clave")) <= 1 Then
             NombreCliente = .fields("Nombre_Completo")
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
   If .RecordCount > 0 And Len(Empresa) > 1 Then
      .MoveFirst
      .Find ("Empresa = '" & Empresa & "' ")
       If Not .EOF Then
          Pict_Version.Cls
         'MsgBox Empresa & vbCrLf & .Fields("Logo_Tipo")
          PosPicX = 10600
          PosLinea = 2000
          LogoTipo1 = Ninguno
          If .fields("Logo_Tipo") <> Ninguno Then
              RutaOrigen = RutaSistema & "\LOGOS\"
              If Existe_File(RutaOrigen & .fields("Logo_Tipo") & ".gif") Then
                 LogoTipo1 = RutaSistema & "\LOGOS\" & .fields("Logo_Tipo") & ".gif"
              Else
                 If Existe_File(RutaOrigen & .fields("Logo_Tipo") & ".jpg") Then LogoTipo1 = RutaSistema & "\LOGOS\" & .fields("Logo_Tipo") & ".jpg"
              End If
          End If
          Pict_Version.Line (PosPicX, PosLinea)-(PosPicX + 2040, PosLinea + 760), Azul, BF
          If LogoTipo1 <> Ninguno Then Pict_Version.PaintPicture LoadPicture(LogoTipo1), PosPicX + 30, PosLinea + 30, 2000, 700
         
         PosPicX = 2100
         PosLinea = 920
         Pict_Version.FontBold = True
         Pict_Version.ForeColor = Azul_Claro
         Pict_Version.FontSize = 9
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print ": " & ULCase(NombreUsuario)
         
         PosLinea = 2000
         PosPicX = 5700
         Pict_Version.FontBold = True
         Pict_Version.ForeColor = Turquesa
         Pict_Version.FontSize = 8
         AltoL = Pict_Version.TextHeight("H")
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Nombre Comercial"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "R.U.C."
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Numero Asignado"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Ciudad"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Dirección"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Teléfono(s)/FAX"
         PosLinea = PosLinea + AltoL
                  
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Representante Legal"
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print "Contador(a)"
         
         PosLinea = 2000
         PosPicX = 7550
         Pict_Version.FontBold = False
         Pict_Version.ForeColor = Azul
        
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Nombre_Comercial")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("RUC")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Item")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print ULCase(.fields("Ciudad"))
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Direccion")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Telefono1") & "-" & .fields("Telefono2") & " / " & .fields("FAX")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Gerente")
         PosLinea = PosLinea + AltoL
         
         Pict_Version.CurrentX = PosPicX
         Pict_Version.CurrentY = PosLinea
         Pict_Version.Print .fields("Contador")

         Dolar = .fields("Cotizacion")
         TextDolar = Format$(Dolar, "#,##0.00")
       End If
   Else
       PosPicX = 6900
       PosLinea = Pict_Version.Top + 2500
       Pict_Version.Cls
       Pict_Version.Picture = LoadPicture(RutaSistema & "\LOGIN.jpg")
       Pict_Version.FontBold = True
       Pict_Version.ForeColor = Verde
       Pict_Version.FontSize = 14
       Pict_Version.CurrentX = PosPicX
       Pict_Version.CurrentY = PosLinea
       AltoL = Pict_Version.TextHeight("H")
       Cadena = "INGRESE SUS CREDENCIALES"
       If Si_No Then Pict_Version.Print Cadena
   End If
  End With
  
 'Version_Sistema
  Si_No = Not (Si_No)
  Pict_Version.FontBold = True
  Pict_Version.FontSize = 10
  AnchoL = Pict_Version.TextWidth(Version_Sistema)
  Pict_Version.CurrentY = Pict_Version.Top + 4700
  Pict_Version.CurrentX = (Pict_Version.width - AnchoL) / 2
  Pict_Version.ForeColor = Amarillo_Claro
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

Private Sub Descargar_FTP_Certificados_Logos()
Dim AdoDBTemp As ADODB.Recordset
Dim ListaDeArchivos As String
Dim rutaFTP As String
Dim Certificados() As String
Dim LogoTipos() As String

Dim sLogoTipos As String
Dim IdC As Byte
Dim iDL As Byte

On Error GoTo error_Handler

  IdC = 0
  sLogoTipos = ""
  
  sSQL = "SELECT Empresa, Ruta_Certificado " _
       & "FROM Empresas " _
       & "WHERE Ruta_Certificado LIKE '%P12' " _
       & "ORDER BY Empresa "
  Select_AdoDB AdoDBTemp, sSQL
  If AdoDBTemp.RecordCount > 0 Then
     Do While Not AdoDBTemp.EOF
        RutaDocumentos = RutaSistema & "\CERTIFIC\" & AdoDBTemp.fields("Ruta_Certificado")
        If Len(Dir$(RutaDocumentos)) = 0 Then
           ReDim Preserve Certificados(IdC) As String
           Certificados(IdC) = AdoDBTemp.fields("Ruta_Certificado")
           IdC = IdC + 1
        End If
        AdoDBTemp.MoveNext
    Loop
  End If
  AdoDBTemp.Close
  
  iDL = 0
  sSQL = "SELECT Logo_Tipo " _
       & "FROM Empresas " _
       & "WHERE Logo_Tipo <> '.' " _
       & "GROUP BY Logo_Tipo " _
       & "ORDER BY Logo_Tipo; "
  Select_AdoDB AdoDBTemp, sSQL
  If AdoDBTemp.RecordCount > 0 Then
     Do While Not AdoDBTemp.EOF
        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.fields("Logo_Tipo") & ".jpg"
        If Len(Dir$(RutaDocumentos)) = 0 Then
           ReDim Preserve LogoTipos(iDL) As String
           LogoTipos(iDL) = AdoDBTemp.fields("Logo_Tipo") & ".jpg"
           iDL = iDL + 1
        End If
        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.fields("Logo_Tipo") & ".gif"
        If Len(Dir$(RutaDocumentos)) = 0 Then
           ReDim Preserve LogoTipos(iDL) As String
           LogoTipos(iDL) = AdoDBTemp.fields("Logo_Tipo") & ".gif"
           iDL = iDL + 1
        End If
        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.fields("Logo_Tipo") & ".png"
        If Len(Dir$(RutaDocumentos)) = 0 Then
           ReDim Preserve LogoTipos(iDL) As String
           LogoTipos(iDL) = AdoDBTemp.fields("Logo_Tipo") & ".gif"
           iDL = iDL + 1
        End If
        AdoDBTemp.MoveNext
    Loop
  End If
  AdoDBTemp.Close
  
  With ftp
      .Inicializar ListEmp
      'Le establecemos la contraseña de la cuenta Ftp
      .Password = ftpPwr
      'Le establecemos el nombre de usuario de la cuenta
      .Usuario = ftpUse
      'Colocamos el puerto de conexion
      .Puerto = 21
      'Establecesmo el nombre del Servidor FTP
      'If InStr(IP_PC.IP_PC, "192.168.27") > 0 Or InStr(IP_PC.IP_PC, "192.168.21") > 0 Then .servidor = "192.168.27.2" Else
      .servidor = ftpSvr
      'conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
      'MsgBox .servidor
       If .ConectarFtp(LstStatud) = False Then
           MsgBox "No se pudo conectar al servidor de Certificados"
           Exit Sub
       End If
       
       LstStatud.Text = LstStatud.Text & .GetDirectorioActual & vbCrLf
       rutaFTP = ""
      'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
       rutaFTP = ""
      'Le indicamos el ListView donde se listarán los archivos
       Set .ListView = LstVwFTP
         
       If IdC > 0 Then
         'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
         .CambiarDirectorio "/SISTEMA/CERTIFIC/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              For J = 0 To UBound(Certificados)
                  If Certificados(J) = LstVwFTP.ListItems(I) Then
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\CERTIFIC\" & LstVwFTP.ListItems(I), True
                     'Exit For
                  End If
              Next J
          Next I
       End If
       If iDL > 0 Then
         'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
         .CambiarDirectorio "/SISTEMA/LOGOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              For J = 0 To UBound(LogoTipos)
                  If UCaseStrg(LogoTipos(J)) = UCaseStrg(LstVwFTP.ListItems(I)) Then
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\LOGOS\" & LstVwFTP.ListItems(I), True
                     'Exit For
                  End If
              Next J
          Next I
       End If
      .Desconectar
   End With
   RatonNormal
Exit Sub
error_Handler:
     MsgBox Err.Description, vbCritical
     RatonNormal
End Sub

