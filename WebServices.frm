VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form FWebServices 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AGENDA ELECTRONICA"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6660
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WebServices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Salir"
      Height          =   435
      Left            =   5145
      TabIndex        =   4
      Top             =   3570
      Width           =   1380
   End
   Begin ComctlLib.TreeView TVXML 
      Height          =   3165
      Left            =   105
      TabIndex        =   3
      Top             =   315
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   5583
      _Version        =   327682
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   735
      Top             =   3570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   210
      Top             =   3570
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Ocultar"
      Height          =   435
      Left            =   3675
      TabIndex        =   1
      Top             =   3570
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   210
      Picture         =   "WebServices.frx":164A
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   2205
      Width           =   1020
   End
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   210
      Top             =   525
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   210
      Top             =   840
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
   Begin MSAdodcLib.Adodc AdoProvincia 
      Height          =   330
      Left            =   210
      Top             =   1155
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
      Caption         =   "Provincia"
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
   Begin MSAdodcLib.Adodc AdoEmp 
      Height          =   330
      Left            =   210
      Top             =   1470
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1785
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BUSQUE EL BENEFICIARIO"
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   6420
   End
End
Attribute VB_Name = "FWebServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XMLDoc As DOMDocument
Dim Primera_Vez As Boolean
Dim ID_XML As Long
Dim ClaveAcceso As String
Dim PatronBusqueda As String
Dim cDestino As String

Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "Agenda Telefonica DiskCover System" & Chr$(0)
    Erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
 
Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Command1_Click()
   CreateIcon
   FWebServices.Visible = False
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Activate()
   'Presentamos datos de la agenda
''    sSQL = "SELECT Cliente,Codigo " _
''         & "FROM Clientes " _
''         & "WHERE Cliente <> '.' " _
''         & "ORDER BY Cliente "
''    SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
''    RatonNormal
''    DCClientes.SetFocus
End Sub

Private Sub Form_Load()
Dim Idx As Integer
Dim ContadorTime As Long
Dim CrearBaseDatos As Boolean
Dim LineaFile As Byte
Dim RutaFile As String
Dim Txt_SMTP_Mails As String

    RatonReloj
    CentrarForm FWebServices
    
    Unidad = Left(CurDir$, 2)
    RutaDestino = Unidad & "\SISTEMA"
    RutaSistema = Unidad & "\SISTEMA"
    RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
    RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
    RutaSysBases = Unidad & "\SYSBASES"
    RutaUpdate = RutaDestino
    ChDir RutaDestino
    
   '---------------------------------------------------------------------------------
   'Datos de Conexion a la Base de Datos en las nubes mysql.diskcoversystem.com:13306
   '---------------------------------------------------------------------------------
    strBaseDatos = "diskcover_empresas"
    strServidor = "mysql.diskcoversystem.com"
    strUsuario = "diskcover"
    strPassword = "disk2017Cover"
    strPuerto = "13306"
    AdoStrCnnMySQL = "DRIVER={MySQL ODBC 3.51 Driver};" _
                   & "SERVER=" & strServidor & ";" _
                   & "DATABASE=" & strBaseDatos & ";" _
                   & "UID=" & strUsuario & ";" _
                   & "PASSWORD=" & strPassword & ";" _
                   & "PORT=" & strPuerto & ";"
   '---------------------------------------------------------------------------------
   'Averiguamos si el MySQL esta en linea
    Si_No = Get_WAN_IP
    
   'Buscamos la cadena de conección a la base en SQL SERVER
    Conectar_Base_Datos
    
    TiempoSistema = Time
    Timer1.Enabled = True
    Timer1.Interval = 1000
    
    UnidadSistema
    MDI_Y_Max = Screen.Height
    MDI_X_Max = Screen.width
    
    EmailRespaldos = Ninguno
    Modulo = "WEBSERVICES"

    Primera_Vez = True
    FechaSistema = Format(date, FormatoFechas)
    
   'Intervalo de espera antes de empezar a sacar los respaldos
    Timer1.Enabled = True
    Timer1.Interval = 1000  '1/2 segundo
    RatonReloj
    
   'Determinar que tipo de bases utilizamos
    Evaluar = False
    SQL_Server = True
    
   'MsgBox Weekday(FechaSistema)
    CodigoUsuario = "ACCESO01"
    NombreUsuario = "Supervisor General"
    Periodo_Contable = "."
    
    RatonReloj
    ConectarAdodc AdoEmp
    ConectarAdodc AdoAux
    ConectarAdodc AdoDetalle
    ConectarAdodc AdoClientes
    ConectarAdodc AdoProvincia
    
    If strWebServices <> "000" And Len(strWebServices) = 3 Then
       Llenar_Empresa_WS
       FWebServices.Caption = "WEB SERVICES"
       CreateIcon
       FWebServices.Visible = False
    Else
       End
    End If
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
         FWebServices.Visible = True
  End Select
End Sub

Private Sub Timer1_Timer()
Dim abData() As Byte
Dim I As Integer

Dim StrgError As String

    RatonReloj
   ' If TiempoSistema = 0 Then TiempoSistema = Time
    Minutos = Minute(Time - TiempoSistema)
    Segundos = Second(Time - TiempoSistema)
    MiTiempo = CSng(Format$(Minutos, "00") & "." & Format$(Segundos, "00"))
    FWebServices.Caption = "WEB-SERVICES: FECHA " & UCase$(Format(date, "dd/MMMM/yyyy")) & " - HORA ACTUAL: " & Format(Time, "HH:MM:SS") & " " & MiTiempo
   '---------------------------------------------------------------------------------------------------------------
   'Averiguamos si el MySQL esta en linea haciendo ping al 8.8.8.8
   '---------------------------------------------------------------------------------------------------------------
    If MiTiempo >= 0.1 Then
       If IP_PC.InterNet And Ping_PC(strServidor) Then
          ID_XML = -1
          sSQL = "SELECT * " _
               & "FROM xml_soap " _
               & "WHERE item = '" & NumEmpresa & "' " _
               & "AND TD = '01' " _
               & "AND opc = 0 " _
               & "ORDER BY fecha DESC " _
               & "LIMIT 1 "
          Select_AdoDB_MySQL AdoRegMySQL, sSQL
          If AdoRegMySQL.RecordCount > 0 Then
            'MsgBox sSQL & vbCrLf & IP_PC.InterNet & vbCrLf & Ping_PC(strServidor)
             ID_XML = AdoRegMySQL.Fields("row_id")
             ClaveAcceso = AdoRegMySQL.Fields("claveAcceso")
             CodigoUsuario = strWebServices & MidStrg(ClaveAcceso, 9, 5)
             TextoFile = AdoRegMySQL.Fields("documento")
             RutaDestino = RutaSysBases & "\TEMP\WS_" & ClaveAcceso & ".xml"
             Escribir_Archivo RutaDestino, TextoFile
          End If
          AdoRegMySQL.Close
          
         'Creamos el arbol del XML si hay archivos
          If ID_XML >= 0 Then
             Encerar_Factura FA
             Set XMLDoc = New DOMDocument
             With XMLDoc
                 .async = False
                 .Load RutaDestino
                  If .parseError.errorCode = 0 Then
                      TVXML.Nodes.Clear
                      If .readyState = 4 Then
                          TVAddNode .documentElement, TVXML
                          If TVXML.Nodes(2) = "infoFactura" Then
                            'MsgBox TVXML.Nodes(2) & vbCrLf & TVXML.Nodes(6)
                             Grabar_WS_CE
                             Kill RutaDestino
                          Else
                          
                          End If
                      End If
                  Else
                     StrgError = .parseError.reason & ", " & .parseError.Line & ", " & .parseError.srcText
                     Control_Procesos Normal, "WS: " & NumEmpresa & ", " & StrgError
                  End If
             End With
          End If
         'Procedemos a borrar los Webservices procesados
          ClaveAcceso = Ninguno
          sSQL = "SELECT * " _
               & "FROM xml_soap " _
               & "WHERE item = '" & NumEmpresa & "' " _
               & "AND TD = '01' " _
               & "AND opc = 1 " _
               & "ORDER BY fecha " _
               & "LIMIT 1 "
          Select_AdoDB_MySQL AdoRegMySQL, sSQL
          If AdoRegMySQL.RecordCount > 0 Then
             ID_XML = AdoRegMySQL.Fields("row_id")
             ClaveAcceso = AdoRegMySQL.Fields("claveAcceso")
          End If
          AdoRegMySQL.Close
          If ClaveAcceso <> Ninguno Then
             FWebServices.Caption = " Eliminando: " & ClaveAcceso
             FWebServices.Refresh
             sSQL = "DELETE " _
                  & "FROM xml_soap " _
                  & "WHERE claveAcceso = '" & ClaveAcceso & "' " _
                  & "AND item = '" & NumEmpresa & "' " _
                  & "AND row_id = " & ID_XML & " " _
                  & "AND opc = 1 "
            Conectar_Ado_Execute_MySQL sSQL
          End If
          TiempoSistema = Time
          RatonNormal
       End If
    End If
   'Fin de actualizacion de el log de ingresos
End Sub

Public Sub Llenar_Empresa_WS()
Dim FechaIniN As Integer
Dim FechaFinN As Integer
Dim NLogoTipo As String
Dim NFirmaDigital As String
Dim NMarcaAgua As String
                   
    PonerLinea = False
    
    sSQL = "SELECT * " _
         & "FROM Modulos " _
         & "WHERE Aplicacion = '" & Modulo & "' "
    SelectAdodc AdoDetalle, sSQL
    If AdoDetalle.Recordset.RecordCount > 0 Then NumModulo = Format(AdoDetalle.Recordset.Fields("Modulo"), "00")
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
    TMail.Adjunto = ""
    TMail.Asunto = ""
    TMail.Mensaje = ""
    TMail.para = ""
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
    
    For I = 0 To 4
        Lista_De_Correos(I).Correo_Electronico = Ninguno
        Lista_De_Correos(I).Contraseña = Ninguno
    Next I
    
    sSQL = "SELECT * " _
         & "FROM Empresas " _
         & "WHERE Item = '" & strWebServices & "' "
    SelectAdodc AdoEmp, sSQL
    With AdoEmp.Recordset
     If .RecordCount > 0 Then
         RatonReloj
         NLogoTipo = .Fields("Logo_Tipo")
         NFirmaDigital = .Fields("Firma_Digital")
         NMarcaAgua = .Fields("Marca_Agua")
         
         LogoTipo = Obtener_File_Grafico(NLogoTipo)
         FirmaDigital = Obtener_File_Grafico(NFirmaDigital)
         MarcaAgua = Obtener_File_Grafico(NMarcaAgua)
         
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Asignacion de correos automáticos para envio a procesos automatizados
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         For I = 0 To 4
             Lista_De_Correos(I).Correo_Electronico = .Fields("Email_Conexion")
             Lista_De_Correos(I).Contraseña = .Fields("Email_Contraseña")
         Next I
         If Len(.Fields("Email_Conexion_CE")) > 1 And Len(.Fields("Email_Contraseña_CE")) > 1 Then
            Lista_De_Correos(4).Correo_Electronico = .Fields("Email_Conexion_CE")
            Lista_De_Correos(4).Contraseña = .Fields("Email_Contraseña_CE")
         End If
         Lista_De_Correos(5).Correo_Electronico = "info@diskcoversystem.com"
         Lista_De_Correos(5).Contraseña = "info2017INFO"
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
         
         SQLDec = ""
         NumEmpresa = .Fields("Item")
         GrupoEmpresa = .Fields("Grupo")
         NumItemTemp = NumEmpresa
         RutaDestino = RutaSistema & "\FONDOS\USUARIOS\" & CodigoUsuario & NumEmpresa & NumModulo & ".gif"
         RutaDocumentos = RutaSysBases & "\CE\CE" & NumEmpresa
         RatonReloj
         Contador = 0
         ContadorRUCCI = 0
         NumItemTemp = NumEmpresa
         OpcCoop = CBool(.Fields("Opc"))
         CentroDeCosto = CBool(.Fields("Centro_Costos"))
         Empresa = .Fields("Empresa")
         EstadoEmpresa = .Fields("Estado")
         Fecha_CE = .Fields("Fecha_CE")
         
         If UCaseStrg(Modulo) <> "UPDATE" Then
            NombreGerente = .Fields("Gerente")
            NombreContador = .Fields("Contador")
            NombreCiudad = .Fields("Ciudad")
            RazonSocial = .Fields("Razon_Social")
            NombreComercial = .Fields("Nombre_Comercial")
            RUC = .Fields("RUC")
            NombrePais = .Fields("Pais")
            CodigoPais = .Fields("CPais")
            CodigoProv = .Fields("CProv")
            
            ReferenciaEmpresa = .Fields("Referencia")
            RUC_Contador = .Fields("RUC_Contador")
            CI_Representante = .Fields("CI_Representante")
            
            FAX = .Fields("FAX")
            Moneda = .Fields("S_M")
            Telefono1 = .Fields("Telefono1")
            Telefono2 = .Fields("Telefono2")
            Direccion = .Fields("Direccion")
            DireccionEstab = .Fields("Direccion")
            Copia_PV = CBool(.Fields("Copia_PV"))
            Mod_Fact = CBool(.Fields("Mod_Fact"))
            Mod_Fecha = CBool(.Fields("Mod_Fecha"))
            Grafico_PV = CBool(.Fields("Grafico_PV"))
            Plazo_Fijo = CBool(.Fields("Plazo_Fijo"))
            ConSucursal = CBool(.Fields("Sucursal"))
            No_Autorizar = CBool(.Fields("No_Autorizar"))
            CodigoDelBanco = .Fields("CodBanco")
            NombreBanco = .Fields("Nombre_Banco")
            Num_Meses_CD = .Fields("Num_CD")
            Num_Meses_CE = .Fields("Num_CE")
            Num_Meses_CI = .Fields("Num_CI")
            Num_Meses_ND = .Fields("Num_ND")
            Num_Meses_NC = .Fields("Num_NC")
            Dec_PVP = .Fields("Dec_PVP")
            Dec_Costo = .Fields("Dec_Costo")
            Dec_IVA = .Fields("Dec_IVA")
            Dec_Cant = .Fields("Dec_Cant")
            
            EmailEmpresa = .Fields("Email")
            EmailContador = .Fields("Email_Contabilidad")
            EmailProcesos = .Fields("Email_Procesos")
            EmailRespaldos = .Fields("Email_Respaldos")
            
            TID_Repres = .Fields("TD")
            Medio_Rol = .Fields("Medio_Rol")
            Sueldo_Basico = .Fields("Sueldo_Basico")
            Cant_Item_PV = .Fields("Cant_Item_PV")
            Cant_Ancho_PV = .Fields("Cant_Ancho_PV")
            Encabezado_PV = .Fields("Encabezado_PV")
            CalcComision = .Fields("Calcular_Comision")
            MascaraCodigoK = .Fields("Formato_Inventario")
            MascaraCodigoA = .Fields("Formato_Activo")
            MascaraCtas = Replace(.Fields("Formato_Cuentas"), "C", "#")
            FormatoCtas = MascaraCtas
            LimpiarCtas = Replace(MascaraCtas, "#", " ")
            ImpCeros = .Fields("Imp_Ceros")
            Fecha_Igualar = .Fields("Fecha_Igualar")
            
           'Documentos Electronicos
            Ambiente = .Fields("Ambiente")
            Obligado_Conta = .Fields("Obligado_Conta")
            ContEspec = .Fields("Codigo_Contribuyente_Especial")
                        
            LimpiarCodigoK = ""
            For I = 1 To Len(MascaraCodigoK)
                If MidStrg(MascaraCodigoK, I, 1) <> "." Then
                   LimpiarCodigoK = LimpiarCodigoK & " "
                Else
                   LimpiarCodigoK = LimpiarCodigoK & "."
                End If
            Next I
            LimpiarCodigoA = ""
            For I = 1 To Len(MascaraCodigoA)
                If MidStrg(MascaraCodigoA, I, 1) <> "." Then
                   LimpiarCodigoA = LimpiarCodigoA & " "
                Else
                   LimpiarCodigoA = LimpiarCodigoA & "."
                End If
            Next I
            
            Porc_Serv = Redondear(.Fields("Servicio") / 100, 2)
            sSQL = "SELECT Porc " _
                  & "FROM Tabla_Por_ICE_IVA " _
                  & "WHERE IVA <> " & Val(adFalse) & " " _
                  & "AND Fecha_Inicio <= #" & BuscarFecha(FechaSistema) & "# " _
                  & "AND Fecha_Final >= #" & BuscarFecha(FechaSistema) & "# " _
                  & "ORDER BY Porc "
            SelectAdodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then Porc_IVA = Redondear(AdoAux.Recordset.Fields("Porc") / 100, 2)
            
            sSQL = "SELECT Descripcion_Rubro " _
                 & "FROM Tabla_Naciones " _
                 & "WHERE CProvincia = '" & CodigoProv & "' " _
                 & "AND TR = 'P' "
            SelectAdodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount > 0 Then NombreProvincia = AdoAux.Recordset.Fields("Descripcion_Rubro")
            
            Informativo_FA = .Fields("LeyendaFA")
            Informativo_FAT = .Fields("LeyendaFAT")
                        
         End If
         NombreRUC = "R.U.C."
        'MsgBox LogoTipo
         Carpeta = .Fields("SubDir")
         EmpresaActual = "[" & RutaEmpresa & "]."
     End If
    End With
    Control_Procesos Normal, "Ingreso a " & Empresa
   '-----------------------------------------------------------------------
   
    Periodo = Ninguno
    EstadoEmpresa = "OK"
    Fecha_CE = FechaSistema
    
   '---------------------------------------------------------------------------------------------------------------
   'Actualizamos las tabla en la web del MySQL
   '---------------------------------------------------------------------------------------------------------------
    If IP_PC.InterNet And Ping_PC(strServidor) Then
       RatonReloj
       Contador = 0
       sSQL = "SELECT ID_Empresa,Estado,Fecha_CE " _
            & "FROM lista_empresas " _
            & "WHERE RUC_CI_NIC = '" & RUC & "' " _
            & "AND Item = '" & NumEmpresa & "' "
       Select_AdoDB_MySQL AdoRegMySQL, sSQL
       If AdoRegMySQL.RecordCount > 0 Then
          IDEntidad = AdoRegMySQL.Fields("ID_Empresa")
          EstadoEmpresa = AdoRegMySQL.Fields("Estado")
          Fecha_CE = AdoRegMySQL.Fields("Fecha_CE")
       End If
       AdoRegMySQL.Close
    End If

   'Actualizamos el estado del internet
    sSQL = "UPDATE Empresas " _
         & "SET Estado = '" & EstadoEmpresa & "' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Estado <> '" & EstadoEmpresa & "' "
    Conectar_Ado_Execute sSQL
       
    sSQL = "UPDATE Empresas " _
         & "SET Fecha_CE = '" & BuscarFecha(Fecha_CE) & "' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Fecha_CE <> '" & BuscarFecha(Fecha_CE) & "' "
    Conectar_Ado_Execute sSQL
    RatonNormal
   'Fin de actualizacion de el log de ingresos
End Sub

Public Sub Grabar_WS_CE()
'AdoEmp
'AdoAux
'AdoAgenda
'AdoClientes
'AdoProvincia

Dim I As Long
Dim Campo As String
Dim Valor As String
Dim Detalle As Boolean
Dim infoAbonos As Boolean
Dim SubTotal As Currency
Dim SubTotalDescuento As Currency
Dim SubTotalIVA As Currency

    TA.Tipo_Pago = "01"
    TA.Abono = 0
    TA.Banco = Ninguno
    TA.Cheque = Ninguno
    TA.Comprobante = Ninguno
    sSQL = "DELETE * " _
         & "FROM Asiento_F " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Conectar_Ado_Execute sSQL
    With TVXML
         Detalle = False
         infoAbonos = False
         Ln_No = 1
         For I = 1 To .Nodes.Count
             FWebServices.Caption = "WEB-SERVICES: FECHA " & UCase$(Format(date, "dd/MMMM/yyyy")) & " - Procesado Documento [" & I & "] "
             Campo = .Nodes(I)
             'MsgBox I & ": " & Campo
             Valor = ""
             If .Nodes(I).children >= 1 Then
                 If Campo <> "detalles" Then I = I + 1
                 Valor = .Nodes(I)
                 If IsNumeric(Valor) And InStr(Valor, ",") Then Valor = CStr(Val(Replace(Valor, ",", "")) / 100)
             End If
             Select Case Campo
               Case "codtda": FA.Cod_CxC = "CXC" & Valor
                              FA.Grupo = Valor
                              FA.SubCta = MidStrg(Valor, 3, 3)
               Case "numint": FA.Observacion = Valor
               Case "serie": FA.Serie = Valor
               Case "numero": FA.Factura = Val(Valor)
               Case "fecemi": FA.Fecha = Valor
               Case "nomcli": FA.Cliente = Valor
               Case "dircli": FA.DireccionC = Valor
               Case "ruccli": FA.RUC_CI = Valor
               Case "telefcli": FA.TelefonoC = Valor
               Case "email": FA.EmailC = Valor
              'Case "moneda"
              'Case "prevta"
              'Case "dscto"
              'Case "valbrut"
              'Case "prevta"
              'Case "igv"
               Case "Total": FA.Total_MN = Val(Valor)
              'Case "detalles"
               Case "detalle"
                    Detalle = True
                    SetAdoAddNew "Asiento_F"
                    Producto = ""
               Case "codart": If Detalle Then SetAdoFields "CODIGO", "03." & Valor
               Case "calidad": Producto = Producto & Valor & ", "
               Case "tallaid": Producto = Producto & Valor
               Case "detallep": Producto = Valor & ", " & Producto
               Case "canti": If Detalle Then SetAdoFields "CANT", Val(Valor)
               Case "dscto"
                    If Detalle Then SubTotalDescuento = Val(Valor) Else FA.Descuento = Val(Valor)
               Case "valven"
                    If Detalle Then
                       Precio = Val(Valor) + SubTotalDescuento
                       SetAdoFields "PRECIO", Precio
                    Else
                       FA.SubTotal = Val(Valor)
                    End If
              'Case "valbrut"
              'Case "valven"
               Case "igv"
                    If Detalle Then SubTotalIVA = Val(Valor) Else FA.Total_IVA = Val(Valor)
               Case "total"
                    If Detalle Then
                        SubTotal = Val(Valor) + SubTotalDescuento
                        SetAdoFields "CODIGO_L", CodigoL
                        SetAdoFields "Cta_SubMod", FA.SubCta
                        SetAdoFields "CodBod", "01"
                        SetAdoFields "CodMar", "01"
                        SetAdoFields "Item", NumEmpresa
                        SetAdoFields "CodigoU", CodigoUsuario
                        SetAdoFields "Cod_Ejec", CodigoUsuario
                        SetAdoFields "A_No", CByte(Ln_No)
                        SetAdoFields "Fecha_V", FA.Fecha
                        SetAdoFields "PRODUCTO", Producto
                        SetAdoFields "PRECIO", Precio
                        SetAdoFields "Total_Desc", SubTotalDescuento
                        SetAdoFields "Total_IVA", SubTotalIVA
                        SetAdoFields "TOTAL", SubTotal
                        SetAdoFields "VALOR_TOTAL", Val(Valor)
                        SetAdoUpdate
                        'MsgBox SubTotal
                        Ln_No = Ln_No + 1
                    End If
               Case "infoAbonos": infoAbonos = True
               Case "numint"
              'Case "moneda"
               Case "tippago": TA.Tipo_Pago = Valor
               Case "tasa": If Val(Valor) > 0 Then TA.Comprobante = "Tasa: " & Valor
               Case "numrefe": TA.Cheque = Valor
               Case "montopag": TA.Abono = Val(Valor)
             End Select
         Next I
    End With
    
   'Verificamos si el Producto existe
    sSQL = "SELECT * " _
         & "FROM Asiento_F " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    SelectAdodc AdoDetalle, sSQL
    With AdoDetalle.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            If .Fields("Total_IVA") > 0 Then BanIVA = True Else BanIVA = False
            sSQL = "SELECT * " _
                 & "FROM Catalogo_Productos " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo_Inv = '" & .Fields("CODIGO") & "' "
            SelectAdodc AdoAux, sSQL
            If AdoAux.Recordset.RecordCount <= 0 Then
               SetAddNew AdoAux
               SetFields AdoAux, "TC", "P"
               SetFields AdoAux, "Codigo_Inv", .Fields("CODIGO")
               SetFields AdoAux, "Producto", .Fields("PRODUCTO")
               SetFields AdoAux, "IVA", BanIVA
               SetFields AdoAux, "Cta_Inventario", "1.1"
               SetFields AdoAux, "Cta_Costo_Venta", "5.1"
               SetFields AdoAux, "Cta_Ventas", "4.1"
               SetFields AdoAux, "Cta_Ventas_0", "4.1"
               SetUpdate AdoAux
            End If
           .MoveNext
         Loop
        .MoveFirst
     End If
    End With
    
   'Verificamos si el cliente existe
    If FA.Cliente = "CONSUMIDOR FINAL" Then
       FA.RUC_CI = "9999999999999"
       FA.EmailC = EmailEmpresa
    End If
    DigVerif = Digito_Verificador(FA.RUC_CI)
    FA.CodigoC = Tipo_RUC_CI.Codigo_RUC_CI
    FA.TD = Tipo_RUC_CI.Tipo_Beneficiario
    sSQL = "SELECT * " _
         & "FROM Clientes " _
         & "WHERE Codigo = '" & FA.CodigoC & "' "
    SelectAdodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount <= 0 Then
       SetAddNew AdoAux
       SetFields AdoAux, "T", Normal
       SetFields AdoAux, "Codigo", FA.CodigoC
       SetFields AdoAux, "Cliente", FA.Cliente
       SetFields AdoAux, "CI_RUC", FA.RUC_CI
       SetFields AdoAux, "Direccion", FA.DireccionC
       SetFields AdoAux, "Telefono", FA.TelefonoC
       SetFields AdoAux, "DirNumero", "SD"
       SetFields AdoAux, "Ciudad", "QUITO"
       SetFields AdoAux, "Email", FA.EmailC
       SetFields AdoAux, "Grupo", FA.Grupo
       SetFields AdoAux, "TD", FA.TD
       SetFields AdoAux, "Prov", "17"
       SetFields AdoAux, "Pais", "593"
       SetFields AdoAux, "FA", adTrue
       SetUpdate AdoAux
    End If
       
    sSQL = "SELECT Codigo, Concepto, Fact, CxC, Serie, Autorizacion, Vencimiento, Fecha, ID " _
         & "FROM Catalogo_Lineas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Codigo = '" & FA.Cod_CxC & "' "
    SelectAdodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       FA.TC = AdoAux.Recordset.Fields("Fact")
       FA.Cta_CxP = AdoAux.Recordset.Fields("CxC")
       FA.Autorizacion = AdoAux.Recordset.Fields("Autorizacion")
       If FA.Serie <> AdoAux.Recordset.Fields("Serie") Then
          AdoAux.Recordset.Fields("Serie") = FA.Serie
          AdoAux.Recordset.Update
       End If
    End If

    sSQL = "SELECT TP, Proceso, Cta_Debe " _
         & "FROM Catalogo_Proceso " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND TP = '" & TA.Tipo_Pago & "' " _
         & "AND DC = 'TA' "
    SelectAdodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       TA.Banco = AdoAux.Recordset.Fields("TP") & " - " & AdoAux.Recordset.Fields("Proceso")
       TA.Cta = AdoAux.Recordset.Fields("Cta_Debe")
    End If
    
    FA.Porc_IVA = Porc_IVA
    FA.Tipo_Pago = "01"
    FA.SP = False
    FA.Fecha_C = FA.Fecha
    FA.Fecha_V = FA.Fecha
    FA.T = Pendiente
    FA.Autorizacion = RUC
    Calculos_Totales_Factura FA
    FA.Saldo_MN = FA.Total_MN
    Grabar_Factura FA, False
    
    TA.TP = FA.TC
    TA.Serie = FA.Serie
    TA.Autorizacion = FA.Autorizacion
    TA.Factura = FA.Factura
    TA.Fecha = FA.Fecha
    TA.CodigoC = FA.CodigoC
    TA.Cta_CxP = FA.Cta_CxP
    If TA.Abono = 0 Then
       TA.Abono = FA.Total_MN
       TA.Tipo_Pago = "01"
       TA.Banco = "01 - EFECTIVO"
       TA.Cheque = Ninguno
       TA.Comprobante = Ninguno
       Grabar_Abonos TA
    ElseIf TA.Abono <= FA.Total_MN Then
       Grabar_Abonos TA
       Diferencia = FA.Total_MN - TA.Abono
       If Diferencia > 0 Then
          TA.Abono = Diferencia
          TA.Tipo_Pago = "01"
          TA.Banco = "01 - EFECTIVO"
          TA.Cheque = Ninguno
          TA.Comprobante = Ninguno
          Grabar_Abonos TA
       End If
    End If
    
    sSQL = "UPDATE xml_soap " _
         & "SET opc = 1 " _
         & "WHERE claveAcceso = '" & ClaveAcceso & "' " _
         & "AND item = '" & NumEmpresa & "' " _
         & "AND row_id = " & ID_XML & " " _
         & "AND opc = 0 "
    Conectar_Ado_Execute_MySQL sSQL
        
   'MsgBox "Succeeded: " & FA.Fecha & " - " & FA.Factura & " - " & CodigoUsuario & vbCrLf & FA.RUC_CI & vbCrLf & sSQL
End Sub

