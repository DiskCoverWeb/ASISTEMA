VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIActivosFijos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   6855
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11400
   Icon            =   "MDIActiv.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIActiv.frx":030A
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
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIActiv.frx":24034C
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIActiv.frx":2409D6
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIActiv.frx":2412B0
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIActiv.frx":2415CA
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIActiv.frx":241EA4
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
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Procesos 
         Caption         =   "De Operacion"
      End
      Begin VB.Menu x1 
         Caption         =   "-"
      End
      Begin VB.Menu MImportExcel 
         Caption         =   "Importar desde Excel"
      End
      Begin VB.Menu x2 
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
   End
End
Attribute VB_Name = "MDIActivosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CambiaEmp_Click()
  RatonReloj
  Modulo = "ACTIVOS FIJOS"
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa MDIActivosFijos, Modulo
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = Me.ScaleHeight - 100
  MDI_X_Max = Me.ScaleWidth - 100
End Sub

Private Sub MDIForm_Load()
  Primera_Vez = True
  Bandera = True
  UnidadSistema
  IngresarClave = True
  Modulo = "ACTIVOS FIJOS"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 10000
  ListEmp.Show 1
  PonerDirEmpresa MDIActivosFijos, Modulo
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Activos Fijos"
  End
End Sub

Private Sub MImportExcel_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub Procesos_Click()
  RatonReloj
  IngActivos.Show
End Sub

Private Sub Reportes_Click()
   RatonReloj
   CatalogoActivos.Show
End Sub

Private Sub SalirS_Click()
  Control_Procesos "Q", "Salir Modulo de Activos Fijos"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict MDIActivosFijos
  Recordar_Tarea_Hora
End Sub

