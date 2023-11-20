VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDICajaChica 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   10575
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13380
   Icon            =   "MDICaja.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDICaja.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13380
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13380
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
      Top             =   10200
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDICaja.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDICaja.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDICaja.frx":2876C
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDICaja.frx":29046
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
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
      Top             =   10080
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu DatosRel 
      Caption         =   "&Archivos"
      Begin VB.Menu DelSyst 
         Caption         =   "Del &Sistema"
         Begin VB.Menu CambClave 
            Caption         =   "Cambio de Clave"
         End
         Begin VB.Menu y1 
            Caption         =   "-"
         End
         Begin VB.Menu MCamboPeriodo 
            Caption         =   "Cambio de Periodo"
         End
      End
      Begin VB.Menu DelOper 
         Caption         =   "De &Operacion"
         Begin VB.Menu MSubMod 
            Caption         =   "Ingresar Catalogo de &Subcuentas"
            Begin VB.Menu IngSubCtasBloq 
               Caption         =   "Ctas. por Cobrar/Ctas. por Pagar"
            End
            Begin VB.Menu MSubIngEgr 
               Caption         =   "Ctas. Ingreso/Egresos/Primas/Centro de Costos"
            End
         End
         Begin VB.Menu MIngGastosCaja 
            Caption         =   "Ingresos/Egresos de Caja C&hica"
            Shortcut        =   ^{F7}
         End
      End
      Begin VB.Menu xxxxx 
         Caption         =   "-"
      End
      Begin VB.Menu MArchivoExcel 
         Caption         =   "Archivos de Excel"
         Shortcut        =   ^I
      End
      Begin VB.Menu Mxx1 
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
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu Mayorizar_Cta 
         Caption         =   "Mayorizar Comprobantes Procesados"
         Shortcut        =   ^T
      End
      Begin VB.Menu CatSubCtas 
         Caption         =   "Catalogo de SubCtas de Bloque"
      End
      Begin VB.Menu MayoresSubCtasBloc 
         Caption         =   "Mayores de SubCuentas"
         Shortcut        =   ^S
      End
      Begin VB.Menu MSep1 
         Caption         =   "-"
      End
      Begin VB.Menu SalBanco 
         Caption         =   "Saldo de Caja/Bancos/Especiales"
      End
      Begin VB.Menu MSaldoCtasEsp 
         Caption         =   "Flujo de Caja Chica"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Herramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu Calculadora 
         Caption         =   "Calculadora"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu Programador 
         Caption         =   "Programador"
      End
      Begin VB.Menu Mdiskcoversystem 
         Caption         =   "www.diskcoversystem.com"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "Ambiente"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDICajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calculadora_Click()
Dim RetVal
  Control_Procesos Normal, "Caculadora"
  RetVal = Shell("CALC.EXE", 1)    ' Ejecuta Calculadora.
End Sub

Private Sub CambClave_Click()
    Titulo = "CAMBIO DE CLAVE"
    Mensajes = "Estimado " & NombreUsuario & ", desea cambiar su clave de acceso?"
    If BoxMensaje = vbYes Then
       RatonReloj
       Control_Procesos Normal, "Cambio de Clave"
       CambClav.Show
    End If
End Sub

Private Sub CatSubCtas_Click()
   Control_Procesos Normal, "Listar Submodulos"
   RatonReloj
   MayorAux1.Show
End Sub

Private Sub ChangeEmp_Click()
  Control_Procesos Normal, "Salir del Sistema"
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub IngSubCtasBloq_Click()
  Control_Procesos Normal, "Catalogo de CxC/CxP"
  RatonReloj
  LCxCxP.Show
End Sub

Private Sub MArchivoExcel_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub MayoresSubCtasBloc_Click()
   Control_Procesos Normal, "Mayores Auxiliares de Submodulos"
   RatonReloj
   MayorAux.Show
End Sub

Private Sub Mayorizar_Cta_Click()
   Mayorizar_Cuentas_SP
   Presenta_Errores_Contabilidad_SP
End Sub

Private Sub MCamboPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
End Sub

Private Sub MDIForm_Activate()
    MDI_X_Max = Screen.width - 150
    MDI_Y_Max = Screen.Height - 1900
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
  Primera_Vez = True
  Bandera = True
  UnidadSistema
  TipoModulo = conta
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "CAJA CHICA"
  MenuDeModulos = True
 'TiempoTarea = Time
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Contabilidad"
  End
End Sub

Private Sub Mdiskcoversystem_Click()
Dim X
  Control_Procesos "Q", "Salir por ingresar a la Pagina WEB"
  MsgBox "Estas a punto de ingresar al Centro de descargas del sistema"
  X = ShellExecute(Me.hwnd, "Open", "http://www.diskcoversystem.com", &O0, &O0, SW_NORMAL)
  End
End Sub

Private Sub MIngGastosCaja_Click()
   Control_Procesos Normal, "I/E: de Caja Chica"
   RatonReloj
   IGastosCaja.Show
End Sub

Private Sub MSaldoCtasEsp_Click()
  Control_Procesos Normal, "Flujo de Caja Chica"
  RatonReloj
  SaldoCtasEspeciales.Show
End Sub

Private Sub MSubIngEgr_Click()
  Control_Procesos Normal, "Catalogo de I/E/CC"
  RatonReloj
  ISubCtas.Show
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


Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
   'Form1.Show
   'Unload FHola
End Sub

Private Sub SalBanco_Click()
  Control_Procesos Normal, "Saldo de Caja Bancos"
  RatonReloj
  SaldoBancos.Show
End Sub

Private Sub SalirSyst1_Click()
  Control_Procesos "Q", "Salir Modulo de Contabilidad"
  End
End Sub

Private Sub Timer1_Timer()
    Ver_Grafico_FormPict
   'If CodigoUsuario = "ACCESO03" Then MAbrirPeriodo.Enabled = True
    If Supervisor = False And Len(CodigoUsuario) > 1 Then
        'MsgBox "No es Supervisor"
        'Seteamos los menus
         NuevoUsu.Enabled = False
        'Respald.Enabled = False
        'RespaldoTotal.Enabled = False
         CierreEjer.Enabled = False
        'MCambioPeriodos.Enabled = False
         Cuentas.Enabled = False
         IngSubCtasBloq.Enabled = False
         MIngGastosCaja.Enabled = False
         MComps.Enabled = False
        'MConciliar.Enabled = False
         MEstFinancieros.Enabled = False
         SalBanco.Enabled = False
         MSaldoCtasEsp.Enabled = False
         LibrosBanco.Enabled = False
         ListMayor.Enabled = False
         ProcBal.Enabled = False
        'EstResultMes.Enabled = False
         MEstResul12M.Enabled = False
        'MayorSubctas.Enabled = False
        'MRepoRete.Enabled = False
         ImpCheques.Enabled = False
         ProcConci.Enabled = False
         
        'Seteamos los menus
         NuevoUsu.Enabled = CNivel(3)
        'Respald.Enabled = CNivel(1) Or CNivel(2) Or CNivel(4) Or CNivel(5)
        'RespaldoTotal.Enabled = CNivel(2)
         CierreEjer.Enabled = CNivel(2)
        'MCambioPeriodos.Enabled = CNivel(2) Or CNivel(3)
         Cuentas.Enabled = CNivel(1) Or CNivel(2) Or CNivel(3)
         IngSubCtasBloq.Enabled = CNivel(1) Or CNivel(2) Or CNivel(3)
         MIngGastosCaja.Enabled = CNivel(2) Or CNivel(3)
         MComps.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(7)
         ProcBal.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5)
        'EstResultMes.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5)
         MEstResul12M.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5)
         
        'MConciliar.Enabled = CNivel(2)
         MEstFinancieros.Enabled = CNivel(2) Or CNivel(7)
         SalBanco.Enabled = CNivel(2)
         MSaldoCtasEsp.Enabled = CNivel(2)
         LibrosBanco.Enabled = CNivel(2) Or CNivel(7)
         ListMayor.Enabled = CNivel(1) Or CNivel(2) Or CNivel(3) Or CNivel(6) Or CNivel(7)
        'MayorSubctas.Enabled = CNivel(2)
        'MRepoRete.Enabled = CNivel(2)
         ImpCheques.Enabled = CNivel(2)
         ProcConci.Enabled = CNivel(2)
    End If
    Recordar_Tarea_Hora
   'Comunicaciones
End Sub

