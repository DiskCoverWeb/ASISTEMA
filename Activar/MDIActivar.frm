VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIActivar 
   BackColor       =   &H80000018&
   Caption         =   "Update"
   ClientHeight    =   5760
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8355
   Icon            =   "MDIActivar.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIActivar.frx":0ECA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   20250
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   20250
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10275
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIActivar.frx":22A98
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIActivar.frx":23122
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIActivar.frx":239FC
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIActivar.frx":23D16
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIActivar.frx":245F0
            Key             =   "Plataforma"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Baltic"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MMantenimiento 
         Caption         =   "Activación del Sistema"
         Shortcut        =   ^A
      End
      Begin VB.Menu MSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "MDIActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MAntesJul_Click()
  RatonReloj
  FMigraAT.Show
End Sub

Private Sub MDIForm_Load()
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
Dim MiArchivo, MiRuta, MiNombre
  ConSubDir = False
  FUnidad.Show 1
  RutaUpdate = RutaDestino & "\SISTEMA"
 'Cadena = InputBox("INGRESE LA UNIDAD DONDE ESTA EL SISTEMA", "UNIDAD DEL SISTEMA", Mid(CurDir$, 1, 3))
  If RutaDestino = "" Then
     MsgBox "No Ingreso la ruta del Sistema"
     End
  Else
     ChDir RutaDestino
     Bandera = True
     SetPrinters.Show 1
     RutaSistema = RutaDestino & "\SISTEMA"
     RutaSysBases = RutaDestino & "\SYSBASES"
     ChDir RutaSistema
    'UnidadSistema
     TipoModulo = Factu
     IngresarClave = True
    'MODULOS
     Modulo = "UPDATE"
     MenuDeModulos = True
     TiempoSistema = Time
     Timer1.Interval = 10000
     IngresarClave = True
   
   RatonReloj
  'Determinar que tipo de bases utilizamos
   Evaluar = False
   SQL_Server = True
   Cadena = Dir(RutaSistema & "\EMPRESA\", vbNormal) 'Recupera la primera entrada.
   Do While Cadena <> ""
      If Cadena <> "." And Cadena <> ".." Then
         If (GetAttr(RutaSistema & "\EMPRESA\" & Cadena) And vbNormal) = vbNormal Then
            If UCase(Cadena) = "DISKCOVE.MDB" Then SQL_Server = False
         End If
      End If
      Cadena = Dir
   Loop
 ' Buscamos la cadena de conección a la base
   If SQL_Server Then
      RutaGeneraFile = RutaSistema & "\SERVER.TXT"
   Else
      RutaGeneraFile = RutaSistema & "\ACCESS.TXT"
   End If
   NumFile = FreeFile
   AdoStrCnn = ""
   Open RutaGeneraFile For Input As #NumFile
   Do While Not EOF(NumFile)
      AdoStrCnn = AdoStrCnn & Input(1, #NumFile) ' Obtiene un carácter.
   Loop
   Close #NumFile
   RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
   RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
 ' Verificamos si la base esta en Microsoft Access o en SQL Server 7.0
   If SQL_Server Then
      PathEmpresa = ""
   Else
      PathEmpresa = UCase(RutaEmpresa & "\DISKCOVE.MDB")
      AdoStrCnn = AdoStrCnn & "Data Source=" & PathEmpresa
   End If
   FechaSistema = Format(date, FormatoFechas)
   NumEmpresa = "000"
   CodigoUsuario = "ACCESO01"
   NombreUsuario = "Supervisor General"
   Empresa = "MODULO DE ACTUALIZACION DE BASES Y DATOS"
   Periodo_Contable = Ninguno
   LogoTipo = UCase(RutaSistema & "\LOGOS\DEFAULT.GIF")
   RatonNormal
   'MsgBox RutaEmpresa
   PonerDirEmpresa MDIActivar, Modulo
   RatonReloj
   FActivar.Show
  End If
End Sub

Private Sub MMantenimiento_Click()
  RatonReloj
  FActivar.Show
End Sub

Private Sub MSalir_Click()
End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict MDIActivar
End Sub


