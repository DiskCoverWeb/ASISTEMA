VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIRespaldo 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8505
   Icon            =   "MDIRespa.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIRespa.frx":1CCA
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
            Picture         =   "MDIRespa.frx":27E88
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIRespa.frx":28512
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIRespa.frx":28DEC
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIRespa.frx":29106
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIRespa.frx":299E0
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
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
      Begin VB.Menu Respald 
         Caption         =   "Respaldo Diario"
         Shortcut        =   ^R
      End
      Begin VB.Menu MProcSuscrip 
         Caption         =   "Procesar Suscripciones"
      End
      Begin VB.Menu MRespaldoTotal 
         Caption         =   "Respaldo Total"
      End
      Begin VB.Menu MCambioPeriodo 
         Caption         =   "Cambio de Perido"
      End
      Begin VB.Menu y 
         Caption         =   "-"
      End
      Begin VB.Menu ChangeEmp 
         Caption         =   "Cambiar de Empresa"
         Shortcut        =   ^E
      End
      Begin VB.Menu SalirSyst1 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "H"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIRespaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.

Private Sub ChangeEmp_Click()
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MCambioPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
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
 ' TipoModulo = Factu
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "RESPALDOS"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 1000
  IngresarClave = True
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Respaldos"
  End
End Sub

Private Sub MProcSuscrip_Click()
  RatonReloj
  Pruebas.Show
  'FSuscripcion.Show
End Sub

Private Sub MRespaldoTotal_Click()
  If ClaveAdministrador Then
     RatonReloj
     RespaldoTotal.Show
  End If
End Sub

Private Sub Respald_Click()
   RatonReloj
   Respaldos.Show
End Sub

Private Sub SalirSyst1_Click()
  Control_Procesos "Q", "Salir Modulo de Respaldos"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
End Sub

