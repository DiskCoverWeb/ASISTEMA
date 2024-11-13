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
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   1164
      ButtonWidth     =   1614
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Monitoreo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Activar"
            Key             =   "Activar"
            Object.ToolTipText     =   "Activa el Monitoreo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Desactivar"
            Key             =   "Desactivar"
            Object.ToolTipText     =   "Desactivar el Monitoreo"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Descargar"
            Key             =   "Descargar"
            Object.ToolTipText     =   "Realizar descargas de Resultados"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Usuarios"
            Key             =   "Usuarios"
            Object.ToolTipText     =   "Presenta Lista de usuarios activos"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Saturado"
            Key             =   "Saturado"
            Object.ToolTipText     =   "CPU Saturado"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Historial"
            Key             =   "Historia_CPU"
            Object.ToolTipText     =   "Historial del CPU Saturado"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tiempo"
            Key             =   "Tiempo_Consulta"
            Object.ToolTipText     =   "Tiempo de Consulta"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Excel"
            Key             =   "Excel"
            Object.ToolTipText     =   "Bajar a excel los resultados"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGMonitoreo 
      Bindings        =   "DiskCover.frx":0ECA
      Height          =   2640
      Left            =   105
      TabIndex        =   3
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
      Left            =   14805
      Picture         =   "DiskCover.frx":0EE5
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   525
      Visible         =   0   'False
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
      Top             =   3885
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
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":1DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":20C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":23E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":26FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":2A17
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":2D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":304B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":3365
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":367F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DiskCover.frx":3999
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Buscando Datos a respaldar, Espere por favor..."
      Height          =   330
      Left            =   525
      TabIndex        =   2
      Top             =   5250
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

Dim TiempoSQL As Single
Dim vTwipsPerPixelX As Single
Dim vTwipsPerPixelY As Single

Dim Opcion_CPU As Byte

Dim LineaConexion(10) As String

Const Opc_CPU_Saturado = 1
Const Opc_Usuarios_En_Linea = 2
Const Opc_CPU_Historia_Saturado = 3
Const Opc_Tiempo_Consulta = 4

Const SQL_CPU_Saturado = "SELECT TOP 10 s.login_name + ':' + s.host_name As login_host_name, s.session_id, r.status, r.cpu_time, r.logical_reads, " _
                       & "r.reads, r.writes, r.total_elapsed_time / (1000 * 60) 'Elaps M', SUBSTRING(st.TEXT, (r.statement_start_offset / 2) + 1, " _
                       & "((CASE r.statement_end_offset WHEN -1 THEN DATALENGTH(st.TEXT) ELSE r.statement_end_offset END - r.statement_start_offset) / 2) + 1) AS statement_text, " _
                       & "COALESCE(QUOTENAME(DB_NAME(st.dbid)) + N'.' + QUOTENAME(OBJECT_SCHEMA_NAME(st.objectid, st.dbid)) " _
                       & "+ N'.' + QUOTENAME(OBJECT_NAME(st.objectid, st.dbid)), '') AS command_text, r.command,  s.program_name, s.last_request_end_time, " _
                       & "s.login_time, r.open_transaction_count " _
                       & "FROM sys.dm_exec_sessions AS s JOIN sys.dm_exec_requests AS r " _
                       & "ON r.session_id = s.session_id CROSS APPLY sys.Dm_exec_sql_text(r.sql_handle) AS st " _
                       & "WHERE r.session_id != @@SPID " _
                       & "ORDER BY r.cpu_time DESC "

Const SQL_Usuarios_En_Linea = "SELECT C.client_net_address, (s.login_name + ':' + s.host_name) As login_host_name, s.program_name, " _
                            & "(CAST(COUNT(s.host_name) As VARCHAR)+ ':' + CAST(SUM(c.session_id) As VARCHAR)) As Procesos_sessions, " _
                            & "MAX(c.connect_time) As Conn_time, (c.net_transport + '-' + c.auth_scheme + ':' + s.client_interface_name) As Protocolo, " _
                            & "MAX(s.login_time) AS Log_Time " _
                            & "FROM sys.dm_exec_connections AS c JOIN sys.dm_exec_sessions AS s ON c.session_id = s.session_id " _
                            & "GROUP BY s.login_name, s.host_name,C.client_net_address,s.program_name, c.net_transport, c.auth_scheme,s.client_interface_name " _
                            & "ORDER BY C.client_net_address, s.login_name DESC, s.host_name, s.program_name "


Const SQL_CPU_Historia_Saturado = "SELECT TOP 15 st.text AS batch_text, SUBSTRING(st.TEXT, (qs.statement_start_offset / 2) + 1,  " _
                                & "((CASE qs.statement_end_offset WHEN - 1 THEN DATALENGTH(st.TEXT) ELSE qs.statement_end_offset END - qs.statement_start_offset) / 2) + 1) AS statement_text, " _
                                & "(qs.total_worker_time / 1000) / qs.execution_count AS avg_cpu_time_ms,(qs.total_elapsed_time / 1000) / qs.execution_count AS avg_elapsed_time_ms, " _
                                & "qs.total_logical_reads / qs.execution_count AS avg_logical_reads,(qs.total_worker_time / 1000) AS cumulative_cpu_time_all_executions_ms, " _
                                & "(qs.total_elapsed_time / 1000) AS cumulative_elapsed_time_all_executions_ms " _
                                & "FROM sys.dm_exec_query_stats qs CROSS APPLY sys.dm_exec_sql_text(sql_handle) st " _
                                & "ORDER BY(qs.total_worker_time / qs.execution_count) DESC "
         
Const MySQL_Tiempo_Consulta = "SELECT login_host_name, program_name, statement_text,Fecha " _
                            & "FROM lista_consulo_cpu " _
                            & "WHERE login_host_name <> '.' " _
                            & "ORDER BY login_host_name, program_name "
                          
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

    TiempoSQL = Time
    
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
    ConectarAdodc AdoMonitoreo
    ConectarAdodc_MySQL AdoMySQL
    
   'MsgBox Weekday(FechaSistema)
    NumEmpresa = "000"
    CodigoUsuario = "ACCESO01"
    NombreUsuario = "Supervisor General"
    Empresa = "MODULO DE MONITOREO"
    Periodo_Contable = "."
    RatonReloj
    
    Tamano_Consultas
    Opcion_CPU = Opc_Usuarios_En_Linea
    Select_Adodc AdoMonitoreo, SQL_Usuarios_En_Linea
    RatonNormal
    FDiskCover.Visible = False
End Sub

Public Sub Consultar_Servidor(sSQLServer As String)
Dim Campo1 As String
Dim Campo2 As String
Dim Campo3 As String
Dim Campo4 As String
Dim InsertSQL As String
Dim InsertReg As String
Dim Insertar As Boolean
    
    Mifecha = BuscarFecha(FechaSistema)
    Insertar = True
    InsertSQL = ""
    Select_Adodc AdoMonitoreo, sSQLServer
    With AdoMonitoreo.Recordset
     If .RecordCount > 0 Then
         DGMonitoreo.Visible = False
         sSQL = "SELECT login_host_name, program_name, statement_text, Fecha " _
              & "FROM lista_consulo_cpu " _
              & "WHERE Fecha = '" & Mifecha & "' " _
              & "ORDER BY program_name "
         Select_Adodc AdoMySQL, sSQL
         Do While Not .EOF
            Select Case sSQLServer
              Case SQL_CPU_Saturado
                   Campo1 = .Fields("program_name")
                   Campo2 = "[Comando: " & .Fields("command") & "]" & " ->> " & .Fields("statement_text")
                   Campo3 = "CPU-Full: " & TrimStrg(MidStrg(.Fields("login_host_name"), 1, 50))
              Case SQL_CPU_Historia_Saturado
                   Campo1 = .Fields("batch_text")
                   Campo2 = .Fields("statement_text")
                   Campo3 = "Historial SQL"
              Case SQL_Usuarios_En_Linea
                   Campo1 = .Fields("client_net_address") & " - " & .Fields("program_name")
                   Campo2 = .Fields("Procesos_sessions") & " - " & .Fields("Protocolo") & " - " & .Fields("Log_Time") & " - " & .Fields("Conn_time")
                   Campo3 = .Fields("login_host_name")
            End Select
            
            Campo1 = Replace(Campo1, "'", "`")
            Campo1 = Replace(Campo1, vbCr, "[CR]")
            Campo1 = Replace(Campo1, vbLf, "[LF]")
            Campo1 = TrimStrg(MidStrg(Campo1, 1, 250))
            
            Campo2 = Replace(Campo2, "'", "`")
            Campo2 = Replace(Campo2, vbCr, "[CR]")
            Campo2 = Replace(Campo2, vbLf, "[LF]")
            
            Campo3 = Replace(Campo3, "'", "`")
            Campo3 = Replace(Campo3, vbCr, "[CR]")
            Campo3 = Replace(Campo3, vbLf, "[LF]")

            If AdoMySQL.Recordset.RecordCount > 0 Then AdoMySQL.Recordset.MoveFirst
            AdoMySQL.Recordset.Find ("program_name = '" & Campo1 & "' ")
            If AdoMySQL.Recordset.EOF Then
               If Insertar Then
                  InsertSQL = "INSERT INTO lista_consulo_cpu(state_type, login_host_name, program_name, statement_text, Fecha) VALUES " & vbCrLf
                  Insertar = False
               End If
               InsertSQL = InsertSQL & "('state_type', '" & Campo3 & "', '" & Campo1 & "', '" & Campo2 & "', '" & Mifecha & "')," & vbCrLf
            End If
           .MoveNext
         Loop
        .MoveFirst
         If InsertSQL <> "" Then
            InsertSQL = MidStrg(InsertSQL, 1, Len(InsertSQL) - 3) & ";"
           'MsgBox InsertSQL
            Conectar_Ado_Execute_MySQL InsertSQL, , "Insertar_MySQL"
         End If
         DGMonitoreo.Visible = True
     End If
    End With
    TiempoSQL = Time
End Sub

Private Sub Timer1_Timer()
  'If MDI_X_Max <> Screen.width - 250 Or MDI_Y_Max <> Screen.Height - 650 Then
  'Tamano_Consultas
  TiempoSistema = Time
  MiTiempo = CSng(Format(Minute(TiempoSistema - TiempoSQL), "00") & "." & Format(Second(TiempoSistema - TiempoSQL), "00"))
  FDiskCover.Caption = Format(TiempoSistema, "HH:MM:SS") & " SERVIDOR: " & strIPServidor _
                     & " [" & Screen.width / vTwipsPerPixelX & " x " & Screen.Height / vTwipsPerPixelY & "]"
  FDiskCover.Refresh
  
  If MiTiempo >= 0.1 Then
    'MsgBox Opcion_CPU & " ............"
     DGMonitoreo.Visible = False
     Select Case Opcion_CPU
       Case Opc_Usuarios_En_Linea
            Consultar_Servidor SQL_Usuarios_En_Linea
       Case Opc_CPU_Saturado
            Consultar_Servidor SQL_CPU_Saturado
       Case Opc_CPU_Historia_Saturado
            Consultar_Servidor SQL_CPU_Historia_Saturado
       Case Opc_Tiempo_Consulta
            Consultar_Servidor MySQL_Tiempo_Consulta
     End Select
     DGMonitoreo.Visible = True
  End If
End Sub

Public Sub Tamano_Consultas()
Dim AltoPantalla As Single
Dim MitadPantalla As Single
            
    'MsgBox Screen.width / vTwipsPerPixelX & " x " & Screen.Height / vTwipsPerPixelY & vbCrLf
    DGMonitoreo.Visible = False
    MDI_X_Max = FDiskCover.width - 150
    MDI_Y_Max = FDiskCover.Height - 150
    
    FDiskCover.Height = MDI_Y_Max
    FDiskCover.Refresh
    
   'Primer Consulta
    DGMonitoreo.Top = 850
    DGMonitoreo.width = (MDI_X_Max - DGMonitoreo.Left) - 40
    DGMonitoreo.Height = (MDI_Y_Max - DGMonitoreo.Top - 550)
                
    DGMonitoreo.Visible = True
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
     ConectarAdodc_MySQL AdoMonitoreo
     sSQL = "SELECT login_host_name, program_name, statement_text, Fecha " _
          & "FROM lista_consulo_cpu " _
          & "WHERE login_host_name <> '.' " _
          & "ORDER BY login_host_name, program_name, Fecha "
     Select_Adodc AdoMonitoreo, sSQL
     With AdoMonitoreo.Recordset
      If .RecordCount > 0 Then
          NumFile = FreeFile
          Open NombreFile For Output As #NumFile
          Print #NumFile, "LISTA DE PROCESOS QUE SATURAN EL SERVIDOR OCASIONANDO LENTITUD"
          Print #NumFile, String(120, "=")
          Do While Not .EOF
             Campo1 = .Fields("login_host_name")
             Campo2 = .Fields("Fecha")
             Campo3 = .Fields("program_name")
             Campo4 = .Fields("statement_text")

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
            .MoveNext
          Loop
          Close #NumFile
      End If
     End With
     RatonNormal
     MsgBox "REVISE EL ARCHIVO SIGUIENTE:" & vbCrLf & NombreFile
     Titulo = "PREGUNTA DE ELIMINACION"
     Mensajes = "VACIAR TABLA TEMPORAL DE CPU SATURADOS?"
     If BoxMensaje = vbYes Then
        RatonReloj
        sSQL = "DELETE FROM lista_consulo_cpu " _
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
    TiempoSQL = Time
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
      Case "Usuarios"
           Opcion_CPU = Opc_Usuarios_En_Linea
           DGMonitoreo.Visible = False
           Select_Adodc AdoMonitoreo, SQL_Usuarios_En_Linea
           DGMonitoreo.Caption = "ESTADO DE SESIONES ABIERTAS POR EQUIPO (" & AdoMonitoreo.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
           'DGMonitoreo.Refresh
           DGMonitoreo.Visible = True
      Case "Saturado"
           Opcion_CPU = Opc_CPU_Saturado
           DGMonitoreo.Visible = False
           Select_Adodc AdoMonitoreo, SQL_CPU_Saturado
           DGMonitoreo.Caption = "ESTADO DE CPU SATURADO (" & AdoMonitoreo.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
           'DGMonitoreo.Refresh
           DGMonitoreo.Visible = True
      Case "Historia_CPU"
           Opcion_CPU = Opc_CPU_Historia_Saturado
           DGMonitoreo.Visible = False
           Select_Adodc AdoMonitoreo, SQL_CPU_Historia_Saturado
           DGMonitoreo.Caption = "ESTADO DE HISTORIA DEL CPU SATURADO (" & AdoMonitoreo.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
           DGMonitoreo.Refresh
           DGMonitoreo.Visible = True
      Case "Tiempo_Consulta"
           Opcion_CPU = Opc_Tiempo_Consulta
           DGMonitoreo.Visible = False
           Select_Adodc AdoMySQL, MySQL_Tiempo_Consulta
           DGMonitoreo.Caption = "ESTADO DE TIEMPOS CONSUMIDOS POR CONSULTAS (" & AdoMonitoreo.Recordset.RecordCount & ") " & Format(Time, "HH:MM:SS")
           'DGMonitoreo.Refresh
           DGMonitoreo.Visible = True
      Case "Excel"
           DGMonitoreo.Visible = False
           ConectarAdodc AdoMonitoreo
           Select_Adodc AdoMonitoreo, SQL_Usuarios_En_Linea
           Exportar_AdoDB_Solo_Excel AdoMonitoreo.Recordset, "USUARIOS_ACTIVOS"

           ConectarAdodc AdoMonitoreo
           Select_Adodc AdoMonitoreo, SQL_CPU_Saturado
           Exportar_AdoDB_Solo_Excel AdoMonitoreo.Recordset, "SERVIDOR_SATURADO"
           
           ConectarAdodc AdoMonitoreo
           Select_Adodc AdoMonitoreo, SQL_CPU_Historia_Saturado
           Exportar_AdoDB_Solo_Excel AdoMonitoreo.Recordset, "HISTORIAL_CONSUMO"
           
           ConectarAdodc_MySQL AdoMonitoreo
           Select_Adodc AdoMonitoreo, MySQL_Tiempo_Consulta
           Exportar_AdoDB_Solo_Excel AdoMonitoreo.Recordset, "TIEMPO_CONSUMIDO"
           DGMonitoreo.Visible = True
    End Select
End Sub

