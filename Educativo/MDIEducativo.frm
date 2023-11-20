VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIEducativo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Instituciones Educativas"
   ClientHeight    =   6360
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8025
   Icon            =   "MDIEducativo.frx":0000
   LinkTopic       =   "MDIReten"
   MouseIcon       =   "MDIEducativo.frx":164A
   Picture         =   "MDIEducativo.frx":1B70
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   8025
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8025
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
      Top             =   5985
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIEducativo.frx":27D2E
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIEducativo.frx":283B8
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIEducativo.frx":28C92
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIEducativo.frx":28FAC
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIEducativo.frx":29886
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
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
      Top             =   5865
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu BaseRel 
      Caption         =   "&Archivos"
      Begin VB.Menu msistema 
         Caption         =   "&Del Sistema"
         Begin VB.Menu MAutNotas 
            Caption         =   "Autorizar Ingreso de Notas"
         End
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
         Begin VB.Menu MAlfabetico 
            Caption         =   "Orden Alfabetico"
         End
         Begin VB.Menu MVerificacion 
            Caption         =   "Verificar Notas"
         End
         Begin VB.Menu MMayaCurr 
            Caption         =   "Maya Curricular"
         End
      End
      Begin VB.Menu MReten 
         Caption         =   "&Ingresar Alumnos"
         Shortcut        =   ^A
      End
      Begin VB.Menu MEstPlant 
         Caption         =   "Estructura del Plantel"
         Shortcut        =   ^R
      End
      Begin VB.Menu MProfMat 
         Caption         =   "Asignar Profesor/Dirigente a Materia/Curso"
         Shortcut        =   ^F
      End
      Begin VB.Menu MImpNotExcel 
         Caption         =   "Importar Notas desde Excel"
         Shortcut        =   ^N
      End
      Begin VB.Menu mCambEmp 
         Caption         =   "Cambiar de &Empresa"
      End
      Begin VB.Menu SalirSystem 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MCatPlantel 
         Caption         =   "Catalogo del Plantel"
      End
      Begin VB.Menu ResPeriodo 
         Caption         =   "Resumen Periodico Materias"
         Visible         =   0   'False
      End
      Begin VB.Menu MImpLib 
         Caption         =   "&Imprimir Libretas/Certificados/Actas de Grado"
         Shortcut        =   ^L
      End
      Begin VB.Menu MReslNotasQuimestres 
         Caption         =   "Resumen Notas/Disciplina/Examen Grados"
         Shortcut        =   ^E
      End
      Begin VB.Menu MListarAlumnosMatriculados 
         Caption         =   "Listar Alumnos Matriculados"
      End
   End
   Begin VB.Menu MHerram 
      Caption         =   "&Herramientas"
      Begin VB.Menu mcalcula 
         Caption         =   "Calculadora"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "h"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIEducativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoAlfa As ADODB.Recordset

Private Sub MAlfabetico_Click()
  RatonReloj
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdoDB AdoAlfa, sSQL
  With AdoAlfa
   If .RecordCount > 0 Then
       Alfabetico = Not (.Fields("Alfabetico"))
   End If
  End With
  AdoAlfa.Close
  sSQL = "UPDATE Catalogo_Periodo_Lectivo "
  Select Case Alfabetico
    Case True: sSQL = sSQL & "SET Alfabetico = " & adTrue & " "
    Case False: sSQL = sSQL & "SET Alfabetico = " & adFalse & " "
  End Select
  sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  If Alfabetico Then
     MDIEducativo.MAlfabetico.Caption = "Listar Alfabéticamente"
  Else
     MDIEducativo.MAlfabetico.Caption = "Listar Por Género"
  End If
  RatonNormal
  MsgBox "Cambio de Estado de listas realizado con éxito"
End Sub

Private Sub MAutNotas_Click()
   If Leer_Campo_Educativo("Cierre_Periodo") Then
      MsgBox "USTED NO PUEDE CAMBIAR DATOS EN UN PERIODO CERRADO"
      Unload FAutorizarNotas
   Else
      RatonReloj
     FAutorizarNotas.Show
   End If
End Sub

Private Sub mcalcula_Click()
Dim RetVal
RetVal = Shell("CALC.EXE", 1)
End Sub

Private Sub mCambEmp_Click()
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

Private Sub MCatPlantel_Click()
  RatonReloj
  FCataEdu.Show
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
  TipoModulo = Cownta
  IngresarClave = True
  
 'MODULOS
  NumModulo = "0"
  Modulo = "EDUCATIVO"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa
  If Alfabetico Then
     MDIFormulario.MAlfabetico.Caption = "Listar Alfabéticamente"
  Else
     MDIFormulario.MAlfabetico.Caption = "Listar Por Género"
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Educativo"
  End
End Sub

Private Sub MEstPlant_Click()
   If Leer_Campo_Educativo("Cierre_Periodo") Then
      MsgBox "USTED NO PUEDE CAMBIAR DATOS EN UN PERIODO CERRADO"
      Unload FAutorizarNotas
   Else
      RatonReloj
      FEducativo.Show
      RatonNormal
   End If
End Sub

Private Sub MImpLib_Click()
  RatonReloj
  FLibretas.Show
End Sub

Private Sub MImpNotExcel_Click()
   If Leer_Campo_Educativo("Cierre_Periodo") Then
      MsgBox "USTED NO PUEDE CAMBIAR DATOS EN UN PERIODO CERRADO"
      Unload FAutorizarNotas
   Else
      RatonReloj
      FImporta.Show
   End If
End Sub

Private Sub MListarAlumnosMatriculados_Click()
  RatonReloj
  FMatriculados.Show
End Sub

Private Sub MMayaCurr_Click()
  If ClaveAdministrador Then
     RatonReloj
     FMalla_Curricular.Show
  End If
End Sub

Private Sub MProfMat_Click()
   If Leer_Campo_Educativo("Cierre_Periodo") Then
      MsgBox "USTED NO PUEDE CAMBIAR DATOS EN UN PERIODO CERRADO"
      Unload FAutorizarNotas
   Else
      RatonReloj
      FProfesorMateria.Show
   End If
End Sub

Private Sub MReslNotasQuimestres_Click()
  RatonReloj
  FNotasQuimestre.Show
End Sub

Private Sub MReten_Click()
  RatonReloj
  FClientesFact.Show
  RatonNormal
End Sub

Private Sub MVerificacion_Click()
  If ClaveAdministrador Then
     RatonReloj
     FVerificacion.Show
  End If
End Sub

Private Sub ResPeriodo_Click()
  RatonReloj
  FResumenNotas.Show
End Sub

Private Sub SalirSystem_Click()
  Control_Procesos "Q", "Salir Modulo de Educativo"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
End Sub
