VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIAnexoSRI 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8505
   Icon            =   "MDIAnexoSRI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIAnexoSRI.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIAnexoSRI.frx":4A5A5
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIAnexoSRI.frx":4AC2F
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIAnexoSRI.frx":4B509
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIAnexoSRI.frx":4B823
            Key             =   "Periodo"
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
   Begin VB.Menu DatosRel 
      Caption         =   "&Archivos"
      Begin VB.Menu DelSyst 
         Caption         =   "Del &Sistema"
         Begin VB.Menu Utilidades 
            Caption         =   "Mantenimiento"
         End
         Begin VB.Menu CambClave 
            Caption         =   "Cambio de Clave"
         End
         Begin VB.Menu NuevoUsu 
            Caption         =   "Ingresar nuevo usuario"
         End
         Begin VB.Menu y 
            Caption         =   "-"
         End
         Begin VB.Menu MCamboPeriodo 
            Caption         =   "Cambio de Periodo"
         End
      End
      Begin VB.Menu DelOper 
         Caption         =   "De &Operacion"
         Begin VB.Menu MCompras 
            Caption         =   "Compras Locales"
         End
      End
      Begin VB.Menu xxxxx 
         Caption         =   "-"
      End
      Begin VB.Menu PRN_LPT1 
         Caption         =   "Establecer y Configurar Impresora"
         Visible         =   0   'False
      End
      Begin VB.Menu ChangeEmp 
         Caption         =   "Cambiar de Empresa"
      End
      Begin VB.Menu SalirSyst1 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MReportes 
      Caption         =   "Reportes"
   End
End
Attribute VB_Name = "MDIAnexoSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
  PonerDirEmpresa MDIAnexoSRI, "ANEXOS SRI"
End Sub


Private Sub MCamboPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa MDIAnexoSRI, Modulo
  End If
End Sub




Private Sub MCompras_Click()
  RatonReloj
  FCompras.Show
End Sub

Private Sub MDIForm_Deactivate()
  Control_Procesos Normal, "Salir del Sistema"
End Sub

 Private Sub MDIForm_Load()
  Primera_Vez = True
  Bandera = True
  SetPrinters.Show 1
  UnidadSistema
  TipoModulo = Conta
  IngresarClave = True
 'MODULOS
  Modulo = "ANEXOS SRI"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa MDIAnexoSRI, Modulo
  Ver_Grafico_FormPict MDIAnexoSRI
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Supervisor = False Then
   ' Seteamos los menus
     NuevoUsu.Enabled = False
     'Respald.Enabled = False
     'RespaldoTotal.Enabled = False
     CierreEjer.Enabled = False
     MCambioPeriodos.Enabled = False
     Cuentas.Enabled = False
     IngSubCtasBloq.Enabled = False
     MIngGastosCaja.Enabled = False
     MComps.Enabled = False
     MConciliar.Enabled = False
     MEstFinancieros.Enabled = False
     SalBanco.Enabled = False
     MSaldoCtasEsp.Enabled = False
     LibrosBanco.Enabled = False
     ListMayor.Enabled = False
     MayorSubctas.Enabled = False
     MRepoRete.Enabled = False
     ImpCheques.Enabled = False
     ProcConci.Enabled = False
   ' Seteamos los menus
     NuevoUsu.Enabled = CNivel_3
     'Respald.Enabled = CNivel_1 Or CNivel_2 Or CNivel_4 Or CNivel_5
     'RespaldoTotal.Enabled = CNivel_2
     CierreEjer.Enabled = CNivel_2
     MCambioPeriodos.Enabled = CNivel_2 Or CNivel_3
     Cuentas.Enabled = CNivel_1 Or CNivel_2 Or CNivel_3
     IngSubCtasBloq.Enabled = CNivel_1 Or CNivel_2 Or CNivel_3
     MIngGastosCaja.Enabled = CNivel_2 Or CNivel_3
     MComps.Enabled = CNivel_2 Or CNivel_3 Or CNivel_4 Or CNivel_5
     MConciliar.Enabled = CNivel_2
     MEstFinancieros.Enabled = CNivel_2
     SalBanco.Enabled = CNivel_2
     MSaldoCtasEsp.Enabled = CNivel_2
     LibrosBanco.Enabled = CNivel_2
     ListMayor.Enabled = CNivel_1 Or CNivel_2 Or CNivel_3 Or CNivel_6
     MayorSubctas.Enabled = CNivel_2
     MRepoRete.Enabled = CNivel_2
     ImpCheques.Enabled = CNivel_2
     ProcConci.Enabled = CNivel_2
  End If
End Sub

Private Sub MDIForm_Terminate()
   Control_Procesos Normal, "Terminación del Sistema"
End Sub
















Private Sub NuevoUsu_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Nuevo Usuario"
     IngClave.Show
  End If
End Sub

Private Sub PRN_LPT1_Click()
    'Establecer CancelError a True
    On Error GoTo ErrHandler
    CommonDialog1.CancelError = False
    'Presentar el cuadro de diálogo Imprimir
    CommonDialog1.Flags = cdlPDPrintSetup
    CommonDialog1.ShowPrinter
    Exit Sub
ErrHandler:
    'El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub SalirSyst1_Click()
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict MDIAnexoSRI
  Recordar_Tarea_Hora
End Sub

Private Sub Utilidades_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Ingreso a Mantenimiento"
     FSeteos.Show
  End If
End Sub

