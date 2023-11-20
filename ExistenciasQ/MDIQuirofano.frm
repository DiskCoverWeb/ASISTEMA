VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIQuirofano 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   7935
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11700
   Icon            =   "MDIQuirofano.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIQuirofano.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11700
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   525
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIQuirofano.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIQuirofano.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIQuirofano.frx":2876C
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIQuirofano.frx":28A86
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIQuirofano.frx":29360
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
      Top             =   7440
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Manten 
         Caption         =   "Del Sistema"
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
      End
      Begin VB.Menu Procesos 
         Caption         =   "De Operacion"
         Begin VB.Menu MSalidaCero 
            Caption         =   "Control de Salida Tarifa Cero"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu MB1 
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
      Begin VB.Menu salinvart 
         Caption         =   "Notas de Entrada/Salida"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "h"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIQuirofano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CambiaEmp_Click()
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
  IngresarClave = True
  NumModulo = "0"
  Modulo = "QUIROFANO"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 10000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Quirofano"
  End
End Sub

Private Sub MSalidaCero_Click()
  RatonReloj
  Empleados = False
  FSalidas_Ceros.Show
End Sub

Private Sub salinvart_Click()
  RatonReloj
  KardexSQLs.Show
End Sub

Private Sub SalirS_Click()
  Control_Procesos "Q", "Salir Modulo de Inventario"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  Recordar_Tarea_Hora
End Sub
