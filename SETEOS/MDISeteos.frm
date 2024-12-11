VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISeteos 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   8220
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11880
   Icon            =   "MDISeteos.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDISeteos.frx":1297D
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   11880
   End
   Begin VB.PictureBox PictBarra 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   105
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7845
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDISeteos.frx":38B3B
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDISeteos.frx":391C5
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDISeteos.frx":39A9F
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISeteos.frx":3A379
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   3
      Top             =   7725
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu DatosRel 
      Caption         =   "&Archivos"
      Begin VB.Menu DelSyst 
         Caption         =   "Del &Sistema"
         Begin VB.Menu Utilidades 
            Caption         =   "Mantenimiento"
            Shortcut        =   ^M
         End
         Begin VB.Menu MMantInstitucion 
            Caption         =   "Mantenimiento Institucion"
         End
         Begin VB.Menu CambClave 
            Caption         =   "Cambio de Clave"
         End
         Begin VB.Menu NuevoUsu 
            Caption         =   "Ingresar nuevo usuario"
         End
         Begin VB.Menu y1 
            Caption         =   "-"
         End
         Begin VB.Menu MCambioPC 
            Caption         =   "Cambio de Periodo Contable"
         End
         Begin VB.Menu MCierreMes 
            Caption         =   "Cierre del Mes"
         End
         Begin VB.Menu MModAud 
            Caption         =   "Modulos de Auditoria"
         End
         Begin VB.Menu MAutorizacionSRI 
            Caption         =   "Autorizaciones del SRI"
         End
         Begin VB.Menu MVerifErrorMayor 
            Caption         =   "Verificar Errores en Mayorización"
         End
      End
      Begin VB.Menu MCopyDatEmp 
         Caption         =   "Copiar Base Datos Empresa"
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu MigrarPlanCta 
         Caption         =   "Subida de Formatos de Excel"
         Shortcut        =   ^E
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu ChangeEmp 
         Caption         =   "Cambiar de Empresa"
      End
      Begin VB.Menu SalirSyst1 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Herramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu Calculadora 
         Caption         =   "Calculadora"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu Programador 
         Caption         =   "Programador"
      End
      Begin VB.Menu MMemos 
         Caption         =   "Memorando"
      End
      Begin VB.Menu MEjemploPDF 
         Caption         =   "Ejemplo PDF"
         Shortcut        =   ^P
      End
      Begin VB.Menu MSalidasxCostos 
         Caption         =   "Salidas por Costos"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "www.diskcoversystem.com"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDISeteos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculadora_Click()
Dim RetVal
  Control_Procesos Normal, "Caculadora"
  RetVal = Shell("CALC.EXE", 1)    ' Ejecuta Calculadora.
End Sub

Private Sub CambClave_Click()
  RatonReloj
  Control_Procesos Normal, "Cambio de Clave"
  CambClav.Show
End Sub

Private Sub ChangeEmp_Click()
  Control_Procesos Normal, "Salir del Sistema"
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MAutorizacionSRI_Click()
If ClaveAdministrador Then
   RatonReloj
   FRenovacion.Show
End If
End Sub

Private Sub MCambioPC_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
End Sub

Private Sub MCierreMes_Click()
  If ClaveContador Then
     RatonReloj
     Cierre.Show
  End If
End Sub

Private Sub MCopyDatEmp_Click()
If ClaveAdministrador Then
   RatonReloj
   FCopyEmpresa.Show
End If
End Sub

Private Sub MDIForm_Activate()
    MDI_X_Max = Screen.width - 150
    MDI_Y_Max = Screen.Height - 1850
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
  Primera_Vez = True
  Bandera = True
  UnidadSistema
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "SETEOS"
  MenuDeModulos = True
  'TiempoTarea = Time
  'TiempoSistema = Time
  'MsgBox TiempoSistema
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Contabilidad"
  End
End Sub

Private Sub MEjemploPDF_Click()
'''Dim pdf As PdfComLib.PdfDoc
'''Dim Archivo As Long
'''Dim NumPag As Integer
'''Dim PDF_ALIGN_CENTER As Integer
'''
''' PDF_ALIGN_CENTER = 2
''' Set pdf = New PdfDoc
''' pdf.AddPage (1)
''' pdf.SetFont "TimesRoman", "BOLD", 10
''' pdf.ClippingText 10, 10, "Esta es otra forma", True
''' pdf.Image "C:\SYSBASES\TEMP\ecuador.jpg", 50, 50, 60, 10, "", 0
''' pdf.SetFont "Helvetica", "", 10
''' pdf.Cell 20, 10, "Hello World !", 0, 0, 1, 0, ""
''' pdf.Text 50, 20, "hola amigos"
''' pdf.Text 50, 30, "hola amigos"
''' pdf.RECT 5, 5, 200, 205, 1
''' pdf.RECT 15, 15, 19.5, 24.5, 1
''' pdf.SaveAsFile ("C:\SYSBASES\TEMP\Ejemplo PDF.pdf")
''' pdf.Closefile
'''
''''Otro Archivo
''' Set pdf = New PdfDoc
''' pdf.OpenPdf
''' pdf.SetFont "Arial", "", 15
''''Page 1
''' pdf.AddPage 0
''' pdf.Bookmark "Page 1", 0, -1
''' pdf.Bookmark "Paragraph 1", 1, -1
''' pdf.Cell 0, 6, "Paragraph 1", 0, 0, 0, 0, ""
''' pdf.ln 50
''' pdf.Bookmark "Paragraph 2", 1, -1
''' pdf.Cell 0, 6, "Paragraph 2", 0, 0, 0, 0, ""
''' pdf.Annotate 60, 30, "First annotation on first page"
''' pdf.Annotate 60, 60, "Second annotation on first page"
''''Page 2
''' pdf.AddPage 0
''' pdf.Bookmark "Page 2", 0, 0
''' pdf.Bookmark "Paragraph 3", 1, -1
''' pdf.Cell 0, 6, "Paragraph 3", 0, 0, 0, 0, ""
''' pdf.Annotate 60, 40, "First annotation on second page"
''' pdf.Annotate 90, 40, "Second annotation on second page"
'''
''' pdf.Bookmark "Paragraph 4", 1, -1
''' pdf.Cell 0, 6, "Paragraph 4", 0, 0, 0, 0, ""
''' pdf.Image "C:\SYSBASES\TEMP\ecuador.jpg", 50, 50, 60, 10, "", 0
'''
''' pdf.SaveAsFile "c:\sysbases\temp\bookmark.pdf"
''' pdf.Closefile
'''
''''Encabezados
'''Set pdf = New PdfDoc
'''pdf.Initialize 1, "C", 3
'''pdf.AliasNbPages "{nb}"
'''pdf.SetFont "Arial", "B", 10
'''
'''pdf.HEADER True
'''pdf.SetXY 1, 1
'''pdf.Cell 210, 10, "This is a Header", 0, 0, 1, 0, ""
'''pdf.SetY 5
'''pdf.HEADER False
'''
'''pdf.FOOTER True
'''pdf.SetY -15
'''pdf.SetFont "Arial", "IB", 8
'''pdf.SetTextColor 128, 100, 128
'''pdf.Cell 0, 10, "Pagina {pg}/{nb}", 0, 0, PDF_ALIGN_CENTER, 0, ""
'''pdf.FOOTER False
'''
'''For NumPag = 1 To 2
'''    pdf.AddPage 1
'''    pdf.SetXY 10, 20
'''    pdf.SetFont "Helvetica", "", 10
'''    pdf.Cell 15, 10, "Body page " & Str(NumPag), 0, 0, 0, 0, ""
'''    pdf.SetFont "Helvetica", "IB", 18
'''    pdf.SetXY 10, 30
'''    pdf.Write 5, "Hola amigos", ""
'''    pdf.SetFont "Times", "B", 16
'''    pdf.Cell 10, 40, "Body page " & Str(NumPag), 0, 0, 0, 0, ""
'''    pdf.SetFont "Times New Roman", "IB", 18
'''    pdf.SetFont "Times", "", 12
'''    pdf.SetXY 100, 50
'''    pdf.SetFillColor Rnd(255) * 255, Rnd(255) * 255, Rnd(255) * 255
'''
'''    pdf.RECT 100, 50, 30, 5, 2
'''    pdf.SetCellMargin 1
'''    pdf.SetDrawColor Rnd(255) * 255, Rnd(255) * 255, Rnd(255) * 255
'''    pdf.RECT 100, 50, 30, 5, 1
'''    pdf.Cell 30, 5, "Otro Texto", 0, 0, 1, 0, ""
'''    pdf.SetXY 100, 60
'''    pdf.Cell 30, 5, "Otro Texto", 0, 0, 1, 0, ""
'''    pdf.SetXY 100, 70
'''    pdf.Cell 30, 5, "Otro Texto", 0, 0, 2, 0, ""
'''    pdf.SetXY 100, 80
'''    pdf.Cell 30, 5, "Otro Texto", 0, 0, 2, 0, ""
'''    pdf.RotatedText 100, 50, "hola estoy rotando 45 grados", 45
'''    pdf.RotatedText 150, 50, "hola estoy rotando 45 grados", 90
'''
'''Next
'''
'''pdf.SaveAsFile ("c:\sysbases\temp\Ejemplo2_PDF.pdf")
'''pdf.Closefile
'''
'''Set pdf = Nothing
'''
'''
'''Archivo = Shell("C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe c:\sysbases\temp\Ejemplo2_PDF.pdf", vbMaximizedFocus)
'''
   FGeneraPDF.Show
End Sub

Private Sub MigrarPlanCta_Click()
  If ClaveContador Then
     RatonReloj
     Control_Procesos Normal, "Subida de Formatos de Master de Excel"
     FImporta.Show
  End If
End Sub

Private Sub MMantInstitucion_Click()
'''  If ClaveAdministrador Then
'''     RatonReloj
'''     FSeteosPlantel.Show
'''  End If
End Sub

Private Sub MMemos_Click()
  RatonReloj
  FMemos.Show
End Sub

Private Sub MModAud_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Modulo Auditoria"
     FAuditoria.Show
  End If
End Sub

Private Sub MSalidasxCostos_Click()
    FGeneraPDF.Show
End Sub

Private Sub MVerifErrorMayor_Click()
   RatonReloj
   Control_Procesos Normal, "Mayorizar Cuentas Erradas"
   MayorizarErrores.Show
End Sub

Private Sub NuevoUsu_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Nuevo Usuario"
     IngClave.Show
  End If
End Sub

Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
   'Form1.Show
   'Unload FHola
End Sub

Private Sub SalirSyst1_Click()
  Control_Procesos "Q", "Salir Modulo de Contabilidad"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  Recordar_Tarea_Hora
  'Comunicacionbes
End Sub

Private Sub Utilidades_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Ingreso a Mantenimiento"
     FSeteos.Show
  End If
End Sub

