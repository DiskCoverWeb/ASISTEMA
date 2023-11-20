VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIHoras 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6525
   Icon            =   "MDIHoras.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIHoras.frx":0742
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   6525
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6525
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
      Top             =   3510
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIHoras.frx":212C9
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIHoras.frx":21953
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIHoras.frx":2222D
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIHoras.frx":22547
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIHoras.frx":22E21
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu MEntSal 
         Caption         =   "Entradas/Salidas"
      End
      Begin VB.Menu MMemos 
         Caption         =   "Memorandos"
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
   Begin VB.Menu MAmbiente 
      Caption         =   "AMBIENTE"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CambiaEmp_Click()
  RatonReloj
  Modulo = "HORAS"
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = MDIFormulario.ScaleHeight - 100
  MDI_X_Max = MDIFormulario.ScaleWidth - 100
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
 'SetPrinters.Show 1
  Primera_Vez = True
  Bandera = True
  UnidadSistema
 ' TipoModulo = conta
  Modulo = "HORAS"
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "CONTABILIDAD"
  MenuDeModulos = True
 'TiempoTarea = Time
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MEntSal_Click()
  RatonReloj
  EntradasSalidas.Show 1
End Sub

Private Sub MMemos_Click()
  RatonReloj
  FMemos.Show
End Sub

Private Sub SalirS_Click()
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
End Sub

