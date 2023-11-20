VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIGastos 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8505
   Icon            =   "MDIGasto.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIGasto.frx":0ECA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8505
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   105
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6075
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIGasto.frx":27088
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIGasto.frx":27712
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIGasto.frx":27FEC
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIGasto.frx":28306
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGasto.frx":28BE0
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBarEstado 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   5955
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu DatosRel 
      Caption         =   "&Archivos"
      Begin VB.Menu MDeOperacion 
         Caption         =   "De Operacion"
         Begin VB.Menu ListCLiProv 
            Caption         =   "Clientes/Proveedores"
         End
         Begin VB.Menu MIngresarRetenciones 
            Caption         =   "Ingresar Retenciones"
            Shortcut        =   ^R
         End
         Begin VB.Menu MAbonosProv 
            Caption         =   "Pago Cheques"
            Shortcut        =   ^P
         End
         Begin VB.Menu MGastosCajaChica 
            Caption         =   "Gastos de Caja Chica"
            Shortcut        =   ^G
         End
         Begin VB.Menu MProcGastosGenerales 
            Caption         =   "Procesar Gastos de Caja Chica"
            Shortcut        =   ^{F5}
         End
      End
      Begin VB.Menu ChangeEmp 
         Caption         =   "Cambiar de Empresa"
      End
      Begin VB.Menu SalirSyst1 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MayoresSubCtasBloc 
         Caption         =   "Mayores de SubCuentas"
         Shortcut        =   ^S
      End
      Begin VB.Menu MSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ListRet 
         Caption         =   "Comprobantes de Retencion"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu ImpCheques 
         Caption         =   "Cheques Procesados"
      End
      Begin VB.Menu MSaldoCtasEsp 
         Caption         =   "Flujo de Caja Chica"
      End
      Begin VB.Menu XXXXX 
         Caption         =   "-"
      End
      Begin VB.Menu MSaldoFactSubMod 
         Caption         =   "Historico de Facturas"
      End
      Begin VB.Menu XXXX 
         Caption         =   "-"
      End
      Begin VB.Menu Programador 
         Caption         =   "Programador"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "AMBIENTE"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''Private Sub Calculadora_Click()
''Dim RetVal
''RetVal = Shell("CALC.EXE", 1)    ' Ejecuta Calculadora.
''End Sub

Private Sub ChangeEmp_Click()
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub ImpCheques_Click()
  RatonReloj
  PCheques.Show
End Sub

Private Sub ListCLiProv_Click()
  RatonReloj
  Nuevo = False
  FClientes.Show
End Sub

Private Sub ListRet_Click()
   RatonReloj
   Retenciones.Show
End Sub

Private Sub MAbonosProv_Click()
  RatonReloj
  AbonosProveedores.Show
End Sub

Private Sub MayoresSubCtasBloc_Click()
   RatonReloj
   MayorAux.Show
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = MDIFormulario.ScaleHeight - 100
  MDI_X_Max = MDIFormulario.ScaleWidth - 100
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
  Primera_Vez = True
  Bandera = True
  UnidadSistema
  TipoModulo = conta
  IngresarClave = True
 'MODULOS
  Modulo = "GASTOS"
  MenuDeModulos = True
  'TiempoTarea = Time
  TiempoSistema = Time
  'MsgBox TiempoSistema
  Timer1.Enabled = True
  Timer1.Interval = 1000
  IngresarClave = True
  ListEmp.Show 1
  'MsgBox Supervisor
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Gerencia"
  End
End Sub

Private Sub MGastosCajaChica_Click()
   Control_Procesos Normal, "I/E: de Caja Chica"
   RatonReloj
   IGastosCaja.Show
End Sub

Private Sub MIngresarRetenciones_Click()
  RatonReloj
  FGastosRet.Show
End Sub

Private Sub MProcGastosGenerales_Click()
  RatonReloj
  ModificarComp = False
  CopiarComp = False
  NuevoComp = True
  FComprobantes.Show
End Sub

Private Sub MSaldoCtasEsp_Click()
  RatonReloj
  SaldoCtasEspeciales.Show
End Sub

Private Sub MSaldoFactSubMod_Click()
  RatonReloj
  HistorialFacturas.Show
End Sub

Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
   'Form1.Show
End Sub

Private Sub SalirSyst1_Click()
  Control_Procesos "Q", "Salir Modulo de Gerencia"
  End
End Sub

Private Sub Timer1_Timer()
  'PresentacionesGIF MDIConta
  Ver_Grafico_FormPict
End Sub

