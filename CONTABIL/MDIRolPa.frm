VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIRolPago 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   6855
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11400
   Icon            =   "MDIRolPa.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIRolPa.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11400
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
      Top             =   6480
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIRolPa.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIRolPa.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIRolPa.frx":2876C
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIRolPa.frx":28A86
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIRolPa.frx":29360
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Baltic"
         Size            =   8,25
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
      Top             =   6360
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu MAnten 
         Caption         =   "Del Sistema"
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
      End
      Begin VB.Menu Procesos 
         Caption         =   "De Operacion"
         Begin VB.Menu MIngEmp 
            Caption         =   "Asignar Empleado al Rol"
            Shortcut        =   ^E
         End
         Begin VB.Menu MActEntSal 
            Caption         =   "Actualizar Entradas/Salidas"
         End
         Begin VB.Menu MAsignarDatosRol 
            Caption         =   "Ingresar Horas Laboradas"
            Shortcut        =   ^H
         End
         Begin VB.Menu MProcRolPagos 
            Caption         =   "Procesar Rol de Pagos"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu x1 
         Caption         =   "-"
      End
      Begin VB.Menu MImpDescExell 
         Caption         =   "Importar Descuento por Excel"
      End
      Begin VB.Menu X4 
         Caption         =   "-"
      End
      Begin VB.Menu MEntSal 
         Caption         =   "Entradas/Salidas"
      End
      Begin VB.Menu MRegistroxTarjetas 
         Caption         =   "Registro por Tarjetas"
      End
      Begin VB.Menu MRolPagoBanco 
         Caption         =   "Enviar Archivo al Banco"
         Begin VB.Menu MBancoInternacional 
            Caption         =   "Banco Internacional"
         End
         Begin VB.Menu MBancoPichincha 
            Caption         =   "Banco del Pichincha"
         End
         Begin VB.Menu MBancoPacifico 
            Caption         =   "Banco del Pacifico"
         End
      End
      Begin VB.Menu X3 
         Caption         =   "-"
      End
      Begin VB.Menu MEnviarMinisterioTrabajo 
         Caption         =   "Enviar Archivo al Ministerio de Trabajo"
      End
      Begin VB.Menu X2 
         Caption         =   "-"
      End
      Begin VB.Menu CambiaEmp 
         Caption         =   "Cambiar de Empresa"
      End
      Begin VB.Menu SalirS 
         Caption         =   "Salir del Sistema"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu MMayorPAgo 
         Caption         =   "Mayorizar Rol de Pagos"
         Shortcut        =   ^T
      End
      Begin VB.Menu MListEmpleados 
         Caption         =   "Listar Empleados"
         Shortcut        =   ^L
      End
      Begin VB.Menu MResumenHorasLab 
         Caption         =   "Resumen Horas Laboradas por Días"
         Shortcut        =   ^D
      End
      Begin VB.Menu MListarCxCxP 
         Caption         =   "Listar SubModulos de CxC/CxP"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "h"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIRolPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CambiaEmp_Click()
  RatonReloj
  Modulo = "ROL PAGOS"
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MActEntSal_Click()
   RatonReloj
   AcEntSal.Show
End Sub

Private Sub MAsignarDatosRol_Click()
  RatonReloj
  HorasEntSal.Show
End Sub

Private Sub MBancoInternacional_Click()
  RatonReloj
  TextoBanco = "INTERNACIONAL"
  FRolPagoBanco.Show
End Sub

Private Sub MBancoPacifico_Click()
  RatonReloj
  TextoBanco = "PACIFICO"
  FRolPagoBanco.Show
End Sub

Private Sub MBancoPichincha_Click()
  RatonReloj
  TextoBanco = "PICHINCHA"
  FRolPagoBanco.Show
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
  IngresarClave = True

 'MODULOS
  NumModulo = "0"
  Modulo = "ROL PAGOS"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Rol de Pagos"
  End
End Sub

Private Sub MEntSal_Click()
  RatonReloj
  EntradasSalidas.Show 1
End Sub

Private Sub MEnviarMinisterioTrabajo_Click()
  RatonReloj
  FRolMinisterioTrabajo.Show
End Sub

Private Sub MImpDescExell_Click()
  If ClaveContador Then
     RatonReloj
     Control_Procesos Normal, "Insertar Descuentos al Rol de Pagos"
     FImporta.Show
  End If
End Sub

Private Sub MIngEmp_Click()
  RatonReloj
  Nuevo = False
  FClientes.Show
End Sub

Private Sub MListarCxCxP_Click()
   Control_Procesos Normal, "Mayores Auxiliares de Submodulos"
   RatonReloj
   MayorAux.Show
End Sub

Private Sub MListEmpleados_Click()
  RatonReloj
  FListarEmpleados.Show
End Sub

Private Sub MMayorPAgo_Click()
   RatonReloj
   Control_Procesos Normal, "Mayorizar Cuentas"
   Mayorizar_Cuentas
End Sub

Private Sub MRegistroxTarjetas_Click()
   RatonReloj
   FTarjetas.Show
End Sub

Private Sub MResumenHorasLab_Click()
  RatonReloj
  ResumenHoras.Show
End Sub

'''Private Sub MSeteos_Click()
'''  If ClaveAdministrador Then
'''     RatonReloj
'''     FSeteos.Show
'''  End If
'''End Sub

Private Sub SalirS_Click()
  Control_Procesos "Q", "Salir Modulo de Rol de Pagos"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  Recordar_Tarea_Hora
End Sub

Private Sub MProcRolPagos_Click()
  RatonReloj
  LRolPagos.Show
End Sub
