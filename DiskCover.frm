VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FDiskCover 
   BackColor       =   &H00800000&
   Caption         =   "RESIDENTE"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15780
   Icon            =   "DiskCover.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   15780
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Activar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Desactivar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Descargar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMonitoreo 
      Bindings        =   "DiskCover.frx":0ECA
      Height          =   2640
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4657
      _Version        =   393216
      BackColor       =   8388608
      BorderStyle     =   0
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.ListBox LstTablas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8715
      TabIndex        =   1
      Top             =   6090
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Timer Timer1 
      Left            =   10920
      Top             =   1785
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   780
      Left            =   210
      Picture         =   "DiskCover.frx":0EE5
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   4830
      Width           =   780
   End
   Begin MSAdodcLib.Adodc AdoMonitoreo 
      Height          =   330
      Left            =   9975
      Top             =   3360
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
      Caption         =   "Monitoreo"
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
   Begin MSAdodcLib.Adodc AdoMySQL 
      Height          =   330
      Left            =   9975
      Top             =   5145
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
      Caption         =   "MySQL"
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
   Begin MSAdodcLib.Adodc AdoMonitoreo1 
      Height          =   330
      Left            =   9975
      Top             =   3780
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
      Caption         =   "Monitoreo"
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
   Begin MSAdodcLib.Adodc AdoMonitoreo2 
      Height          =   330
      Left            =   9975
      Top             =   4200
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
      Caption         =   "Monitoreo"
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
   Begin MSAdodcLib.Adodc AdoMonitoreo3 
      Height          =   330
      Left            =   9975
      Top             =   4620
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
      Caption         =   "Monitoreo"
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
   Begin MSDataGridLib.DataGrid DGMonitoreo1 
      Bindings        =   "DiskCover.frx":1DAF
      Height          =   2640
      Left            =   105
      TabIndex        =   4
      Top             =   3465
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4657
      _Version        =   393216
      BackColor       =   8388608
      BorderStyle     =   0
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin MSDataGridLib.DataGrid DGMonitoreo2 
      Bindings        =   "DiskCover.frx":1DCB
      Height          =   2640
      Left            =   4515
      TabIndex        =   5
      Top             =   840
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4657
      _Version        =   393216
      BackColor       =   8388608
      BorderStyle     =   0
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin MSDataGridLib.DataGrid DGMonitoreo3 
      Bindings        =   "DiskCover.frx":1DE7
      Height          =   2640
      Left            =   4515
      TabIndex        =   6
      Top             =   3675
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4657
      _Version        =   393216
      BackColor       =   8388608
      BorderStyle     =   0
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9030
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":1E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":211D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":2437
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":2751
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Buscando Datos a respaldar, Espere por favor..."
      Height          =   330
      Left            =   4305
      TabIndex        =   2
      Top             =   6090
      Visible         =   0   'False
      Width           =   4320
   End
End
Attribute VB_Name = "FDiskCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LineaConexion(10) As String
Dim TiempoSQL1 As Single
Dim TiempoSQL2 As Single
Dim TiempoSQL3 As Single
Dim TiempoSQL4 As Single
Dim vTwipsPerPixelX As Single
Dim vTwipsPerPixelY As Single

'---------------------------------------------------------------------------------
'Datos de Conexion a la Base de Datos en las nubes mysql.diskcoversystem.com:13306
'---------------------------------------------------------------------------------
Const AdoStrCnnMySQL = "DRIVER={MySQL ODBC 5.1 Driver};" _
                     & "SERVER=db.diskcoversystem.com;" _
                     & "DATABASE=diskcover_empresas;" _
                     & "USER=diskcover;" _
                     & "PASSWORD=disk2017@Cover;" _
                     & "PORT=13306;" _
                     & "OPTION=3;"

Private Sub Form_Load()
Dim Idx As Integer
Dim ContadorTime As Long
Dim CrearBaseDatos As Boolean
Dim LineaFile As Byte
Dim RutaFile As String
Dim Txt_SMTP_Mails As String

   '---------------------------------------------------------------------------------
   'Datos de Conexion a la Base de Datos en las nubes mysql.diskcoversystem.com:13306
   '---------------------------------------------------------------------------------
    RatonReloj
    
    vTwipsPerPixelX = Screen.TwipsPerPixelX
    vTwipsPerPixelY = Screen.TwipsPerPixelY
    
    FDiskCover.width = Screen.width
    FDiskCover.Height = Screen.Height - 550

    TiempoSQL1 = Time
    TiempoSQL2 = TiempoSQL1
    TiempoSQL3 = TiempoSQL1
    TiempoSQL4 = TiempoSQL1
    
   'Formulario Principal
    FDiskCover.Top = 0
    FDiskCover.Left = 0
    
    Timer1.Enabled = True
    Timer1.Interval = 1000  '5 segundos
    
    CreateIcon
            
    NombreComercial = "PRISMANET PROFESIONAL S.A."
    RazonSocial = "DISKCOVER SYSTEM"
    Telefono1 = "593-2-6052430"
    Telefono2 = "593-2-999654196"
    Direccion = "Atacames N23-226 y Av. La Gasca"
    NombreCiudad = "Quito"
    NombrePais = "Ecuador"
        
    EmailRespaldos = Ninguno
    Modulo = "DISKCOVER"

    Primera_Vez = True
    FechaSistema = Format(date, FormatoFechas)
    FechaRespaldo = UCase(Format(Year(FechaSistema), "0000") & "-" & Mid(MesesLetras(Month(FechaSistema)), 1, 3) & "-" & Format(Day(FechaSistema), "00"))
    For Idx = 0 To 4
        LineaConexion(Idx) = ""
    Next Idx
    RatonReloj
    Unidad = Left(CurDir$, 2)
    RutaDestino = Unidad & "\SISTEMA"
    RutaSistema = Unidad & "\SISTEMA"
    RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
    RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
    RutaSysBases = Unidad & "\SYSBASES"
    RutaUpdate = RutaDestino
    ChDir RutaDestino
   'Determinar que tipo de bases utilizamos
    Evaluar = False
    SQL_Server = True
    Conectar_Base_Datos
    
   'MsgBox Weekday(FechaSistema)
    NumEmpresa = "000"
    CodigoUsuario = "ACCESO01"
    NombreUsuario = "Supervisor General"
    Empresa = "MODULO DE MONITOREO"
    Periodo_Contable = "."
    RatonReloj
    ConectarAdodc AdoMonitoreo
    ConectarAdodc AdoMonitoreo1
    ConectarAdodc AdoMonitoreo2
    ConectarAdodc_MySQL AdoMonitoreo3
    ConectarAdodc_MySQL AdoMySQL
    Tamano_Consultas
    Consultar_Usuarios_En_Linea
    Consultar_Tiempo_Consulta
    Consultar_CPU_Saturado
    Consultar_CPU_Historia_Saturado
    RatonNormal
    FDiskCover.Visible = False
End Sub

Private Sub Timer1_Timer()
  'If MDI_X_Max <> Screen.width - 250 Or MDI_Y_Max <> Screen.Height - 650 Then
  'Tamano_Consultas
  TiempoSistema = Time
  FDiskCover.Caption = Format(TiempoSistema, "HH:MM:SS") & " SERVIDOR: " & strIPServidor _
                     & " [" & Screen.width / vTwipsPerPixelX & " x " & Screen.Height / vTwipsPerPixelY & "]"
  FDiskCover.Refresh
  
  MiTiempo = TiempoSQL4
  MiTiempo = CSng(Format(Minute(TiempoSistema - MiTiempo), "00") & "." & Format(Second(TiempoSistema - MiTiempo), "00"))
  If MiTiempo >= 5 Then
     Consultar_Tiempo_Consulta
     TiempoSQL4 = TiempoSistema
     FDiskCover.Caption = Format(TiempoSistema, "HH:MM:SS") & " SERVIDOR: " & strIPServidor & " - Consultar_Tiempo_Consulta"
     FDiskCover.Refresh
  End If
  
  MiTiempo = TiempoSQL1
  MiTiempo = CSng(Format(Minute(TiempoSistema - MiTiempo), "00") & "." & Format(Second(TiempoSistema - MiTiempo), "00"))
  If MiTiempo >= 0.03 Then
     Consultar_Usuarios_En_Linea
     TiempoSQL1 = TiempoSistema
  End If
  
  MiTiempo = TiempoSQL2
  MiTiempo = CSng(Format(Minute(TiempoSistema - MiTiempo), "00") & "." & Format(Second(TiempoSistema - MiTiempo), "00"))
  If MiTiempo >= 0.1 Then
     Consultar_CPU_Saturado
     TiempoSQL2 = TiempoSistema
  End If

  MiTiempo = TiempoSQL3
  MiTiempo = CSng(Format(Minute(TiempoSistema - MiTiempo), "00") & "." & Format(Second(TiempoSistema - MiTiempo), "00"))
  If MiTiempo >= 1 Then
     Consultar_CPU_Historia_Saturado
     TiempoSQL3 = TiempoSistema
     FDiskCover.Caption = Format(TiempoSistema, "HH:MM:SS") & " SERVIDOR: " & strIPServidor & " - Consultar_CPU_Historia_Saturado"
     FDiskCover.Refresh
  End If

End Sub

Public Sub Tamano_Consultas()
Dim AltoPantalla As Single
Dim MitadPantalla As Single
            
    'MsgBox Screen.width / vTwipsPerPixelX & " x " & Screen.Height / vTwipsPerPixelY & vbCrLf
        
    MDI_X_Max = FDiskCover.width - 150
    MDI_Y_Max = FDiskCover.Height - 150
    
    FDiskCover.Height = MDI_Y_Max
    FDiskCover.Refresh
    
   'Primer Consulta
    AltoPantalla = ((MDI_Y_Max - DGMonitoreo.Top) / 3) - 260
    MitadPantalla = ((MDI_X_Max - DGMonitoreo.Left) / 2) - 40
    DGMonitoreo.Top = 550
    DGMonitoreo.width = MitadPantalla
    DGMonitoreo.Height = (MDI_Y_Max - DGMonitoreo.Top - 550)
    
   'Segunda Consulta
    DGMonitoreo1.width = MitadPantalla - 100
    DGMonitoreo1.Height = AltoPantalla
    DGMonitoreo1.Top = DGMonitoreo.Top
    DGMonitoreo1.Left = DGMonitoreo.Left + MitadPantalla + 100

   'Tercera Consulta
    DGMonitoreo2.width = MitadPantalla - 100
    DGMonitoreo2.Height = AltoPantalla
    DGMonitoreo2.Left = DGMonitoreo.Left + MitadPantalla + 100
    DGMonitoreo2.Top = DGMonitoreo1.Height + DGMonitoreo1.Top + 100
    
   'Cuarta Consulta
    DGMonitoreo3.width = MitadPantalla - 100
    DGMonitoreo3.Height = AltoPantalla
    DGMonitoreo3.Left = DGMonitoreo.Left + MitadPantalla + 100
    DGMonitoreo3.Top = DGMonitoreo2.Top + DGMonitoreo2.Height + 100
            
    DGMonitoreo.Visible = True
    DGMonitoreo1.Visible = True
    DGMonitoreo2.Visible = True
    DGMonitoreo3.Visible = True
End Sub

Public Sub Consultar_Usuarios_En_Linea()
's.client_interface_name,
    DGMonitoreo.Visible = False
    sSQL = "SELECT C.client_net_address, (s.login_name + ':' + s.host_name) As login_host_name, s.program_name, (CAST(COUNT(s.host_name) As VARCHAR)+ ':' + CAST(SUM(c.session_id) As VARCHAR)) As Procesos_sessions, " _
         & "MAX(c.connect_time) As Conn_time, (c.net_transport + '-' + c.auth_scheme + ':' + s.client_interface_name) As Protocolo, MAX(s.login_time) AS Log_Time " _
         & "FROM sys.dm_exec_connections AS c JOIN sys.dm_exec_sessions AS s ON c.session_id = s.session_id " _
         & "GROUP BY s.login_name, s.host_name,C.client_net_address,s.program_name, c.net_transport, c.auth_scheme,s.client_interface_name " _
         & "ORDER BY C.client_net_address, s.login_name DESC, s.host_name, s.program_name "
    Select_Adodc AdoMonitoreo, sSQL
    DGMonitoreo.Visible = True
    DGMonitoreo.Caption = "ESTADO DE SESIONES ABIERTAS POR EQUIPO (" & AdoMonitoreo.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
    'MsgBox "0:" & Format(Time, "HH:MM:SS")
End Sub

Public Sub Consultar_CPU_Saturado()
Dim Campo1 As String
Dim Campo2 As String

   'Presenta el consumo de Consultas SQL cuando el CPU esta saturado
    DGMonitoreo1.Visible = False
    sSQL = "SELECT TOP 10 s.login_name + ':' + s.host_name As login_host_name, s.session_id, r.status, r.cpu_time, r.logical_reads, " _
         & "r.reads, r.writes, r.total_elapsed_time / (1000 * 60) 'Elaps M', SUBSTRING(st.TEXT, (r.statement_start_offset / 2) + 1, " _
         & "((CASE r.statement_end_offset WHEN -1 THEN DATALENGTH(st.TEXT) ELSE r.statement_end_offset END - r.statement_start_offset) / 2) + 1) AS statement_text, " _
         & "COALESCE(QUOTENAME(DB_NAME(st.dbid)) + N'.' + QUOTENAME(OBJECT_SCHEMA_NAME(st.objectid, st.dbid)) " _
         & "+ N'.' + QUOTENAME(OBJECT_NAME(st.objectid, st.dbid)), '') AS command_text, r.command,  s.program_name, s.last_request_end_time, " _
         & "s.login_time, r.open_transaction_count " _
         & "FROM sys.dm_exec_sessions AS s JOIN sys.dm_exec_requests AS r " _
         & "ON r.session_id = s.session_id CROSS APPLY sys.Dm_exec_sql_text(r.sql_handle) AS st " _
         & "WHERE r.session_id != @@SPID " _
         & "ORDER BY r.cpu_time DESC "
    Select_Adodc AdoMonitoreo1, sSQL
    DGMonitoreo1.Caption = "ESTADO DE CPU SATURADO (" & AdoMonitoreo1.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
    With AdoMonitoreo1.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             Campo1 = TrimStrg(MidStrg(Replace(.Fields("program_name"), "'", """"), 1, 250))
             Campo2 = "[Comando: " & .Fields("command") & "]" & vbCrLf & .Fields("statement_text")
            'MsgBox "..."
             sSQL = "SELECT login_host_name, program_name, statement_text, Fecha " _
                  & "FROM lista_consulo_cpu " _
                  & "WHERE program_name = '" & Campo1 & "' "
             Select_Adodc AdoMySQL, sSQL
             If AdoMySQL.Recordset.RecordCount <= 0 Then
                AdoMySQL.Recordset.AddNew
                AdoMySQL.Recordset.Fields("login_host_name") = "CPU-Full: " & TrimStrg(MidStrg(.Fields("login_host_name"), 1, 50))
                AdoMySQL.Recordset.Fields("program_name") = Campo1
                AdoMySQL.Recordset.Fields("statement_text") = Campo2
                AdoMySQL.Recordset.Fields("Fecha") = FechaSistema
                AdoMySQL.Recordset.Update
             End If
            .MoveNext
          Loop
         .MoveFirst
      End If
    End With
    DGMonitoreo1.Visible = True
End Sub

Public Sub Consultar_CPU_Historia_Saturado()
Dim Campo1 As String
Dim Campo2 As String

   'Presenta los ultimos 12 registros que consumen demaciado CPU
    DGMonitoreo2.Visible = False
    sSQL = "SELECT TOP 15 st.text AS batch_text, SUBSTRING(st.TEXT, (qs.statement_start_offset / 2) + 1,  " _
         & "((CASE qs.statement_end_offset WHEN - 1 THEN DATALENGTH(st.TEXT) ELSE qs.statement_end_offset END - qs.statement_start_offset) / 2) + 1) AS statement_text, " _
         & "(qs.total_worker_time / 1000) / qs.execution_count AS avg_cpu_time_ms,(qs.total_elapsed_time / 1000) / qs.execution_count AS avg_elapsed_time_ms, " _
         & "qs.total_logical_reads / qs.execution_count AS avg_logical_reads,(qs.total_worker_time / 1000) AS cumulative_cpu_time_all_executions_ms, " _
         & "(qs.total_elapsed_time / 1000) AS cumulative_elapsed_time_all_executions_ms " _
         & "FROM sys.dm_exec_query_stats qs CROSS APPLY sys.dm_exec_sql_text(sql_handle) st " _
         & "ORDER BY(qs.total_worker_time / qs.execution_count) DESC "
    Select_Adodc AdoMonitoreo2, sSQL
    DGMonitoreo2.Caption = "ESTADO DE HISTORIA DEL CPU SATURADO (" & AdoMonitoreo2.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
    With AdoMonitoreo2.Recordset
        ' MsgBox .RecordCount
      If .RecordCount > 0 Then
          Do While Not .EOF
             Campo1 = TrimStrg(MidStrg(Replace(.Fields("batch_text"), "'", """"), 1, 250))
             Campo2 = .Fields("statement_text")
             
             sSQL = "SELECT login_host_name, program_name, statement_text, Fecha " _
                  & "FROM lista_consulo_cpu " _
                  & "WHERE program_name = '" & Campo1 & "' "
             Select_Adodc AdoMySQL, sSQL
             If AdoMySQL.Recordset.RecordCount <= 0 Then
                'MsgBox "...." & BuscarFecha(FechaSistema)
                AdoMySQL.Recordset.AddNew
                AdoMySQL.Recordset.Fields("login_host_name") = "Historial SQL"
                AdoMySQL.Recordset.Fields("program_name") = Campo1
                AdoMySQL.Recordset.Fields("statement_text") = Campo2
                AdoMySQL.Recordset.Fields("Fecha") = FechaSistema
                AdoMySQL.Recordset.Update
             End If
            .MoveNext
          Loop
         .MoveFirst
      End If
    End With
    DGMonitoreo2.Visible = True
    'MsgBox "3:" & Format(Time, "HH:MM:SS")
End Sub

Public Sub Consultar_Tiempo_Consulta()
   'Captura el tiempo total de CPU empleado por una consulta junto con el plan de consulta y las ejecuciones totales
    DGMonitoreo3.Visible = False
    sSQL = "SELECT login_host_name, program_name, statement_text,Fecha " _
         & "FROM lista_consulo_cpu " _
         & "WHERE login_host_name <> '.' " _
         & "ORDER BY login_host_name, program_name "
    Select_Adodc AdoMonitoreo3, sSQL
    DGMonitoreo3.Visible = True
    DGMonitoreo3.Caption = "ESTADO DE TIEMPOS CONSUMIDOS POR CONSULTAS (" & AdoMonitoreo3.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
End Sub

Public Sub Descargar_Reporte()
Dim NumFile As Long
Dim NombreFile As String
Dim Campo1 As String
Dim Campo2 As String
Dim Campo3 As String
Dim Campo4 As String

    NombreFile = RutaSysBases & "\TEMP\lista_consulo_cpu_" & Replace(FechaSistema, "/", "-") & ".xml"
  If Len(NombreFile) > 1 Then
     RatonReloj
     sSQL = "SELECT login_host_name, program_name, statement_text, Fecha " _
          & "FROM lista_consulo_cpu " _
          & "WHERE login_host_name <> '.' " _
          & "ORDER BY login_host_name, program_name, Fecha "
     Select_Adodc AdoMySQL, sSQL
     If AdoMySQL.Recordset.RecordCount > 0 Then
          NumFile = FreeFile
          Open NombreFile For Output As #NumFile
          Print #NumFile, "LISTA DE PROCESOS QUE SATURAN EL SERVIDOR OCASIONANDO LENTITUD"
          Print #NumFile, String(120, "=")
          Do While Not AdoMySQL.Recordset.EOF
             Campo1 = AdoMySQL.Recordset.Fields("login_host_name")
             Campo2 = AdoMySQL.Recordset.Fields("Fecha")
             Campo3 = AdoMySQL.Recordset.Fields("program_name")
             Campo4 = AdoMySQL.Recordset.Fields("statement_text")

             Print #NumFile, Campo1 & String(12, vbTab) & "FECHA DE INSIDENCIA: " & Campo2
             Print #NumFile, String(120, "=")
             Print #NumFile, "ENCABEZADO DEL PROCESO:"
             Print #NumFile, Campo3
             Print #NumFile, String(120, "-")
             Print #NumFile, "PROGRAMACION QUE OCASIONA LA SATURACION:"
             Print #NumFile, String(40, "-")
             Print #NumFile, Campo4
             Print #NumFile, String(120, "_")
             Print #NumFile, String(120, "=") & vbCrLf
             AdoMySQL.Recordset.MoveNext
          Loop
          Close #NumFile
     End If
     RatonNormal
     MsgBox "REVISE EL ARCHIVO SIGUIENTE:" & vbCrLf & NombreFile
     Titulo = "PREGUNTA DE ELIMINACION"
     Mensajes = "VACIAR TABLA TEMPORAL DE CPU SATURADOS?"
     If BoxMensaje = vbYes Then
        RatonReloj
        sSQL = "DELETE " _
             & "FROM lista_consulo_cpu " _
             & "WHERE login_host_name <> '.' "
        Conectar_Ado_Execute_MySQL sSQL
        RatonNormal
     End If
  End If
End Sub

Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbsize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "Conexiones de DiskCover System" & Chr$(0)
    Erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
 
Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbsize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  X = X / Screen.TwipsPerPixelX
  Select Case X
    Case WM_LBUTTONDOWN
         Caption = "Left Click"
    Case WM_RBUTTONDOWN
         Caption = "Right Click"
    Case WM_MOUSEMOVE
         Caption = "Move"
    Case WM_LBUTTONDBLCLK
         Caption = "Double Click"
         FDiskCover.Visible = True
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           End
      Case "Activar"
           CreateIcon
           FDiskCover.Visible = False
      Case "Desactivar"
           DeleteIcon
      Case "Descargar"
           Descargar_Reporte
    End Select
End Sub

