VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIAnexos 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   8505
   Icon            =   "MDIAnexo.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIAnexo.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10830
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIAnexo.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIAnexo.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIAnexo.frx":2876C
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIAnexo.frx":28A86
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAnexo.frx":29360
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
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   19080
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   19080
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
         Begin VB.Menu MModAud 
            Caption         =   "Modulos de Auditoria"
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
         Begin VB.Menu MCliente 
            Caption         =   "Clientes o Proveedores"
            Shortcut        =   ^P
         End
         Begin VB.Menu MComprasLocales 
            Caption         =   "Compras Locales"
            Shortcut        =   ^A
         End
         Begin VB.Menu MVentasLocales 
            Caption         =   "Ventas Locales"
            Shortcut        =   ^N
         End
         Begin VB.Menu MExportaciones 
            Caption         =   "Exportaciones"
            Shortcut        =   ^E
         End
         Begin VB.Menu MImportaciones 
            Caption         =   "Importaciones"
            Shortcut        =   ^I
         End
         Begin VB.Menu MRDEP 
            Caption         =   "Relacion de Dependencia"
            Shortcut        =   {F7}
         End
         Begin VB.Menu MAnulados 
            Caption         =   "Anulados"
            Shortcut        =   {F6}
         End
         Begin VB.Menu MRecap 
            Caption         =   "Recap"
         End
         Begin VB.Menu cal 
            Caption         =   "Calculos"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MCambiaremp 
         Caption         =   "Cambiar de Empresa"
      End
      Begin VB.Menu MnSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MAnexoTrans 
         Caption         =   "Generar Anexo Transaccional"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MGeneral 
         Caption         =   "Códigos de Retención (AIR)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MGeneralPorcen 
         Caption         =   "Resumen Retenciones IVA y Fuente"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MVerificacionError 
         Caption         =   "Corrección de Formularios"
         Begin VB.Menu MVerfCompras 
            Caption         =   "Compras"
         End
         Begin VB.Menu MVerifVentas 
            Caption         =   "Ventas"
         End
         Begin VB.Menu MVerifExportaciones 
            Caption         =   "Exportaciones"
         End
         Begin VB.Menu MVerifImportaciones 
            Caption         =   "Importaciones"
         End
      End
      Begin VB.Menu MCatalogRet 
         Caption         =   "Catalogo de Codigos de Retencion"
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
   End
End
Attribute VB_Name = "MDIAnexos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cal_Click()
FCalculos.Show
End Sub

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

Private Sub MAnexoTrans_Click()
   RatonReloj
   FAnexoTransaccional.Show
End Sub

Private Sub MAnulados_Click()
  RatonReloj
  FAnulados.Show
End Sub

Private Sub MCambiaremp_Click()
  Control_Procesos Normal, "Salir del Sistema"
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa MDIAnexos, Modulo
End Sub

Private Sub MCamboPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa MDIAnexos, Modulo
  End If
End Sub

Private Sub MCatalogRet_Click()
  RatonReloj
  FCodigosRetencion.Show
End Sub

Private Sub MCliente_Click()
    RatonReloj
    FClientes.Show
End Sub

Private Sub MComprasLocales_Click()
    RatonReloj
    Nuevo = True
    Trans_No = 20
    FCompras.Show
End Sub

Private Sub MDIForm_Deactivate()
  Control_Procesos Normal, "Salir del Sistema"
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = Me.ScaleHeight - 100
  MDI_X_Max = Me.ScaleWidth - 100
End Sub

Private Sub MDIForm_Load()
  Primera_Vez = True
  Bandera = True
  SetPrinters.Show 1
  UnidadSistema
  TipoModulo = conta
  IngresarClave = True
 'MODULOS
  Modulo = "ANEXOS SRI"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa MDIAnexos, Modulo
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

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Anexos SRI"
  End
End Sub

Private Sub MExportaciones_Click()
  RatonReloj
  Nuevo = True
  Trans_No = 20
  FExportaciones.Show
End Sub

Private Sub MGeneral_Click()
  RatonReloj
  FReportes.Show
End Sub

Private Sub MGeneralPorcen_Click()
    RatonReloj
    'FReportesG.Show
    FLibroRetenciones.Show
End Sub

Private Sub MImportaciones_Click()
   RatonReloj
   Nuevo = True
   Trans_No = 20
   FImportaciones.Show
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

Private Sub MnSalir_Click()
  Control_Procesos "Q", "Salir Modulo de Anexos SRI"
  End
End Sub

Private Sub MRDEP_Click()
  RatonReloj
  FRetencion.Show
End Sub

Private Sub MRecap_Click()
   RatonReloj
   Nuevo = True
   FRecap.Show
End Sub

Private Sub MVentasLocales_Click()
   RatonReloj
   Nuevo = True
   Trans_No = 20
   FVentas.Show
End Sub

Private Sub MVerfCompras_Click()
    RatonReloj
    Nuevo = False
    Trans_No = 21
    FCompras.Show
End Sub

Private Sub MVerifExportaciones_Click()
    RatonReloj
    Nuevo = False
    Trans_No = 21
    FExportaciones.Show
End Sub

Private Sub MVerifImportaciones_Click()
    RatonReloj
    Nuevo = False
    Trans_No = 21
    FImportaciones.Show
End Sub

Private Sub MVerifVentas_Click()
    RatonReloj
    Nuevo = False
    Trans_No = 21
    FVentas.Show
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

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict MDIAnexos
  Recordar_Tarea_Hora
End Sub

Private Sub Utilidades_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Ingreso a Mantenimiento"
     FSeteos.Show
  End If
End Sub

