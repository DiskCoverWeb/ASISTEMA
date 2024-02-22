VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIConta 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   10575
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13380
   Icon            =   "MDIConta.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIConta.frx":164A
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
            Picture         =   "MDIConta.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIConta.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIConta.frx":2876C
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIConta.frx":29046
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
         Begin VB.Menu NuevoUsu 
            Caption         =   "Ingresar nuevo usuario"
         End
         Begin VB.Menu CambClave 
            Caption         =   "Cambio de Clave"
         End
         Begin VB.Menu y1 
            Caption         =   "-"
         End
         Begin VB.Menu MModAud 
            Caption         =   "Modulos de Auditoria"
         End
         Begin VB.Menu MAutorizacionSRI 
            Caption         =   "Autorizaciones del SRI"
         End
         Begin VB.Menu y 
            Caption         =   "-"
         End
         Begin VB.Menu ProcesosContable 
            Caption         =   "Cierre del Mes"
         End
         Begin VB.Menu CierreEjer 
            Caption         =   "Cierre del Ejercicio"
         End
         Begin VB.Menu MCamboPeriodo 
            Caption         =   "Cambio de Periodo"
         End
      End
      Begin VB.Menu DelOper 
         Caption         =   "De &Operacion"
         Begin VB.Menu xxxxxxx 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu Cuentas 
            Caption         =   "Ingresar &Catalogo de Cuentas"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu MSubMod 
            Caption         =   "Ingresar Catalogo de &Subcuentas"
            Begin VB.Menu IngSubCtasBloq 
               Caption         =   "Ctas. por Cobrar/Ctas. por Pagar"
            End
            Begin VB.Menu MSubIngEgr 
               Caption         =   "Ctas. Ingreso/Egresos/Primas/Centro de Costos"
            End
         End
         Begin VB.Menu MIngBenef 
            Caption         =   "Ingresar Clientes/Proveedores"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu MSubCtaProyectos 
            Caption         =   "Ingresar SubCuentas de Proyectos"
         End
         Begin VB.Menu B13 
            Caption         =   "-"
         End
         Begin VB.Menu MComps 
            Caption         =   "&Ingresar Comprobantes"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu R12 
            Caption         =   "-"
         End
         Begin VB.Menu MIngGastosCaja 
            Caption         =   "Ingresos/Egresos de Caja C&hica"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu MModPrimas 
            Caption         =   "Modificacion de Primas"
         End
         Begin VB.Menu MCostoDelProyecto 
            Caption         =   "Contabilizacion de Costos de Proyectos"
         End
         Begin VB.Menu xxxxxx 
            Caption         =   "-"
         End
         Begin VB.Menu ConcBank 
            Caption         =   "Conciliación &Bancaria de Debitos/Creditos"
         End
         Begin VB.Menu RecibirXML 
            Caption         =   "Recepcion de Archivos Autorizados (SRI)"
            Shortcut        =   ^X
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
      Begin VB.Menu MPagoProgBancos 
         Caption         =   "Pagos Programados Bancos"
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
      Begin VB.Menu ListCtas 
         Caption         =   "Catalogo de Cuentas"
      End
      Begin VB.Menu CatSubCtas 
         Caption         =   "Catalogo de SubCtas de Bloque"
      End
      Begin VB.Menu MCatGasCaja 
         Caption         =   "Cuotas Pendientes de Prestamos"
      End
      Begin VB.Menu MCatRolPagos 
         Caption         =   "Catalogo de Rol de Pagos"
      End
      Begin VB.Menu MCatalogoRetencion 
         Caption         =   "Catalogo de Codigos de Retencion"
      End
      Begin VB.Menu MBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu ListDiaGen 
         Caption         =   "Diario General"
         Shortcut        =   ^D
      End
      Begin VB.Menu LibrosBanco 
         Caption         =   "Libro Banco"
         Shortcut        =   ^B
      End
      Begin VB.Menu ListMayor 
         Caption         =   "Mayores Auxiliares"
         Shortcut        =   ^M
      End
      Begin VB.Menu MMayorAuxConcepto 
         Caption         =   "Mayores Auxiliares por Conceptos"
      End
      Begin VB.Menu MayoresSubCtasBloc 
         Caption         =   "Mayores de SubCuentas"
         Shortcut        =   ^S
      End
      Begin VB.Menu MSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MListComprobantes 
         Caption         =   "Comprobantes Procesados"
         Shortcut        =   ^L
      End
      Begin VB.Menu ListRet 
         Caption         =   "Comprobantes de Retencion"
         Shortcut        =   ^R
      End
      Begin VB.Menu MComprobantesLC 
         Caption         =   "Comprobantes de Liquidacion de Compras"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ImpCheques 
         Caption         =   "Cheques Procesados"
      End
      Begin VB.Menu ProcConci 
         Caption         =   "Conciliacion Bancaria"
      End
      Begin VB.Menu MImpComp 
         Caption         =   "Imprimir Lista de Comprobantes"
      End
      Begin VB.Menu MBarAT 
         Caption         =   "-"
      End
      Begin VB.Menu SalBanco 
         Caption         =   "Saldo de Caja/Bancos/Especiales"
      End
      Begin VB.Menu MSaldoFactSubMod 
         Caption         =   "Saldo de Facturas en SubModulos"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MSaldoCtasEsp 
         Caption         =   "Flujo de Caja Chica"
         Shortcut        =   ^O
      End
      Begin VB.Menu MRepSubCtasCosto 
         Caption         =   "Reporte de SubCtas de Costos"
         Begin VB.Menu MIngFactCostos 
            Caption         =   "Ingresar Facturas de Costos"
         End
         Begin VB.Menu ListSubCtaBloc 
            Caption         =   "Reporte de Costos"
         End
      End
      Begin VB.Menu MRepCoop 
         Caption         =   "-"
      End
      Begin VB.Menu MReportesUAF 
         Caption         =   "Informes a la UAF (Cooperativas)"
      End
      Begin VB.Menu xxxx 
         Caption         =   "-"
      End
      Begin VB.Menu BuscarDatos 
         Caption         =   "Buscar Datos"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MAnexosTrans 
      Caption         =   "Anexos Transaccionales"
      Begin VB.Menu MGenerarAT 
         Caption         =   "Generar Anexos Transaccionales"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MCodigosAir 
         Caption         =   "Codigos de Retencion (AIR)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MResumenAT 
         Caption         =   "Resumen de Retenciones"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MRelacionDependencia 
         Caption         =   "Relación por Dependencia"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MAnulados 
         Caption         =   "Anulados"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu MEstFinancieros 
      Caption         =   "&Estados Financieros"
      Begin VB.Menu Mayorizar 
         Caption         =   "&Mayorizar Comprobantes Procesados"
         Shortcut        =   ^T
      End
      Begin VB.Menu MMayUltCompProc 
         Caption         =   "Mayorizar &Ultimos Comprobantes Procesados"
         Visible         =   0   'False
      End
      Begin VB.Menu zzz 
         Caption         =   "-"
      End
      Begin VB.Menu ProcBal 
         Caption         =   "Balances de Comprobación/Situación/General"
         Shortcut        =   ^G
      End
      Begin VB.Menu xx 
         Caption         =   "-"
      End
      Begin VB.Menu MEstResul12M 
         Caption         =   "Resumen Analitico de Utilidad/Perdida"
      End
      Begin VB.Menu xxx 
         Caption         =   "-"
      End
      Begin VB.Menu MSalVencSubCtas 
         Caption         =   "Balances de SubModulos"
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
      Begin VB.Menu MReindexa_Cuentas 
         Caption         =   "Reindexar Cuentas"
      End
      Begin VB.Menu MMemos 
         Caption         =   "Memorando"
         Shortcut        =   ^E
      End
      Begin VB.Menu MConexionOracle 
         Caption         =   "Conexion Oracle"
      End
      Begin VB.Menu Mdiskcoversystem 
         Caption         =   "www.diskcoversystem.com"
      End
      Begin VB.Menu MEmails 
         Caption         =   "Enviar Correo Electronico"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "Ambiwnte"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BuscarDatos_Click()
  Control_Procesos Normal, "Buscar Datos"
  RatonReloj
  BuscarContabil.Show
End Sub

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

Private Sub CierreEjer_Click()
  If ClaveGerente Then
     RatonReloj
     Control_Procesos Normal, "Cierre del Ejercicio"
     CierreEjercicio.Show
  End If
End Sub

Private Sub ConcBank_Click()
  Control_Procesos Normal, "Conciliar Depositos/Retiros"
  RatonReloj
  Conciliacion.Show
End Sub

Private Sub Cuentas_Click()
  Control_Procesos Normal, "Catalogo de Cuentas"
  RatonReloj
  FCatalogo_Cuentas.Show
End Sub

Private Sub ChangeEmp_Click()
  Control_Procesos Normal, "Salir del Sistema"
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub ImpCheques_Click()
  Control_Procesos Normal, "Cheques Procesados"
  RatonReloj
  PCheques.Show
End Sub

'''Private Sub IngDC_Click()
'''  Control_Procesos Normal, "Conciliar NC/ND"
'''  RatonReloj
'''  NotaDebCred.Show
'''End Sub

Private Sub IngSubCtasBloq_Click()
  Control_Procesos Normal, "Catalogo de CxC/CxP"
  RatonReloj
  LCxCxP.Show
End Sub

Private Sub LibrosBanco_Click()
   Control_Procesos Normal, "Libro Banco"
   RatonReloj
   LibroBanco.Show
End Sub

'''Private Sub ListCLiProv_Click()
'''  Control_Procesos Normal, "Listar Clientes/Proveedores"
'''  RatonReloj
'''  Nuevo = False
'''  ListClientes.Show
'''End Sub

Private Sub ListCtas_Click()
   Control_Procesos Normal, "Listar Catalogo de Cuentas"
   RatonReloj
   CatalogoCtas.Show
End Sub

Private Sub ListDiaGen_Click()
  Control_Procesos Normal, "Diario General"
  RatonReloj
  LibroDiario.Show
End Sub

Private Sub ListMayor_Click()
  Control_Procesos Normal, "Mayores Auxiliares"
  RatonReloj
  Individual = False
  ListMayorizacion1.Show
End Sub

Private Sub ListRet_Click()
   Control_Procesos Normal, "Listar Retenciones"
   RatonReloj
   Retenciones.Show
End Sub

Private Sub ListSubCtaBloc_Click()
  RatonReloj
  RSubCtas.Show
End Sub

Private Sub MAnulados_Click()
  RatonReloj
  FAnulados.Show
End Sub

Private Sub MArchivoExcel_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub MAutorizacionSRI_Click()
If ClaveAdministrador Then
   RatonReloj
   FRenovacion.Show
End If
End Sub

Private Sub MayoresSubCtasBloc_Click()
   Control_Procesos Normal, "Mayores Auxiliares de Submodulos"
   RatonReloj
   MayorAux.Show
End Sub

Private Sub Mayorizar_Click()
   Mayorizar_Cuentas_SP
   Presenta_Errores_Contabilidad_SP
  'Mayorizar_Cuentas
End Sub

''''Private Sub MAyuda_Click()
''''  Contador = 0
''''  Cadena = "" & vbCrLf
''''  For Each Control In MDIConta.Controls
''''      Cadena = Cadena & Control.Name & vbTab & vbTab
''''      Contador = Contador + 1
''''      If Contador >= 5 Then
''''         Cadena = Cadena & vbCrLf
''''         Contador = 0
''''      End If
''''  Next Control
''''  'MsgBox Cadena
''''  ''FHola.Show
''''''Dim RV As Long
''''''  Cadena = RutaSistema & "\HELPS\" & UCaseStrg(App.EXEName) & ".hlp"
''''''  RV = WinHelp(MDIConta.hWnd, Cadena, &H3, CLng(0))
''''End Sub

Private Sub MCamboPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
End Sub

Private Sub MCatalogoRetencion_Click()
  RatonReloj
  FCodigosRetencion.Show
End Sub

Private Sub MCatGasCaja_Click()
  Control_Procesos Normal, "Cartera de Prestamos"
  RatonReloj
  MayorAux2.Show
End Sub

Private Sub MCatRolPagos_Click()
   Control_Procesos Normal, "Listar Catalogo Rol de Pagos"
   RatonReloj
   FListarEmpleados.Show
End Sub

Private Sub MCodigosAir_Click()
  RatonReloj
  FReportes.Show
End Sub

Private Sub MComprobantesLC_Click()
   Control_Procesos Normal, "Listar Retenciones"
   RatonReloj
   FLiquidacionCompras.Show
End Sub

Private Sub MComps_Click()
  RatonReloj
  Control_Procesos Normal, "Procesar Comprobantes Contables"
  NuevoComp = True
  ModificarComp = False
  CopiarComp = False
  Co.CodigoB = Ninguno
  Co.Numero = 0
  Co.TP = CompDiario
  FComprobantes.Show
End Sub

Private Sub MConexionOracle_Click()
   'FOracle.Show
   FGeneraPDF.Show
End Sub

Private Sub MCostoDelProyecto_Click()
  Control_Procesos Normal, "Generacion de Costos del Proyecto"
  RatonReloj
  FCostosDelProyecto.Show
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
''  FEsperar.Show
''  MsgBox ".."
''  Unload FEsperar
''  MsgBox "..."
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

Private Sub MEmails_Click()
''' Cadena = " HOLA AMIGOS MIOS, ESTA ES UNA PRUENA DE CADENA "
''' MsgBox "         1         2         3         4         5" & vbCrLf _
'''      & "12345678901234567890123456789012345678901234567890" & vbCrLf _
'''      & Cadena & vbCrLf _
'''      & "Left: '" & LeftStrg(Cadena, 10) & "'" & vbCrLf _
'''      & "Right: '" & RightStrg(Cadena, 10) & "'" & vbCrLf _
'''      & "Mid 10,10: '" & MidStrg(Cadena, 10, 10) & "'" & vbCrLf _
'''      & "InStr 'MIOS': '" & InStr(Cadena, "MIOS") & "'"
'''
  FEnviarMails.Show
''  FGeneraPDF.Show
''  Exportar_Datagrid_Execel "d:\Nomina1.xls"
''  Exportar_Datagrid_Execel "d:\Nomina2.xls"
''  TMail.ListaError = ""
''  For I = 1 To 10
''    TMail.Adjunto = ""
''    TMail.Asunto = "Prueba de mail"
''    TMail.Mensaje = "Este es un mensaje de prueba de correo enviado por medio del programa DiskCover System"
''    TMail.para = "diskcover@msn.com"
''    FEnviarCorreos.Show 1
''  Next I
End Sub

Private Sub MEstResul12M_Click()
  Control_Procesos Normal, "Resumen Analitico de U/P"
  RatonReloj
  EstadoResult12Meses.Show
End Sub

Private Sub MGenerarAT_Click()
  RatonReloj
  FAnexoTransaccional.Show
End Sub

Private Sub MImpComp_Click()
  Control_Procesos Normal, "Imprimir Comprobates por Lotes"
  RatonReloj
  ImprimirComprobantes.Show
End Sub

Private Sub MIngBenef_Click()
  Control_Procesos Normal, "Registro de Clientes o Proveedores"
  RatonReloj
  Nuevo = False
  FClientes.Show
End Sub

Private Sub MIngFactCostos_Click()
  RatonReloj
  IngFactCostos.Show
End Sub

Private Sub MIngGastosCaja_Click()
   Control_Procesos Normal, "I/E: de Caja Chica"
   RatonReloj
   IGastosCaja.Show
End Sub

Private Sub MListComprobantes_Click()
   Control_Procesos Normal, "Listar Comprobantes Procesados"
   RatonReloj
   NumeroComp = 0
   FechaComp = Ninguno
   FListComprobantes.Show
End Sub

Private Sub MMayorAuxConcepto_Click()
  Control_Procesos Normal, "Mayores Auxiliares por Concepto"
  RatonReloj
  Individual = True
  ListMayorizacion1.Show
End Sub

Private Sub MMayUltCompProc_Click()
   RatonReloj
   Mayorizar2.Show
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

Private Sub MModPrimas_Click()
   Control_Procesos Normal, "Modificacion de primas"
   RatonReloj
   FPrimas.Show
End Sub

Private Sub MPagoProgBancos_Click()
  RatonReloj
  FPagosBancos.Show
End Sub

Private Sub MReindexa_Cuentas_Click()
    RatonReloj
    Parametros = "'" & NumEmpresa & "','" & Periodo_Contable & "' "
    Ejecutar_SP "sp_Reindexar_Periodo", Parametros
    
    Mayorizar_Cuentas_SP
    Presenta_Errores_Contabilidad_SP
    RatonNormal
End Sub

Private Sub MRelacionDependencia_Click()
  RatonReloj
  FRetencion.Show
End Sub

Private Sub MReportesUAF_Click()
  Control_Procesos Normal, "Buscar Datos"
  RatonReloj
  FReportesUAF.Show
End Sub

Private Sub MResumenAT_Click()
  RatonReloj
  FLibroRetenciones.Show
End Sub

Private Sub MSaldoCtasEsp_Click()
  Control_Procesos Normal, "Flujo de Caja Chica"
  RatonReloj
  SaldoCtasEspeciales.Show
End Sub

Private Sub MSaldoFactSubMod_Click()
  Control_Procesos Normal, "Listar Facturas de CxC/CxP"
  RatonReloj
  SaldoSubCtasVence.Show
End Sub

Private Sub MSalVencSubCtas_Click()
  Control_Procesos Normal, "Balances de Submodulos"
  RatonReloj
  BalanceSubCtas.Show
End Sub

Private Sub MSubCtaProyectos_Click()
  RatonReloj
  FCatalogo_Costos.Show
End Sub

Private Sub MSubIngEgr_Click()
  Control_Procesos Normal, "Catalogo de I/E/CC"
  RatonReloj
  ISubCtas.Show
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
    On Error GoTo errHandler
    CommonDialog1.CancelError = False
    'Presentar el cuadro de diálogo Imprimir
    CommonDialog1.Flags = cdlPDPrintSetup
    CommonDialog1.ShowPrinter
    Exit Sub
errHandler:
    'El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub ProcBal_Click()
  Control_Procesos Normal, "Listar/Procesar Balances Financieros"
  RatonReloj
  BalanceComp.Show
End Sub

Private Sub ProcConci_Click()
  Control_Procesos Normal, "Listar Conciliacion Bancaria"
  RatonReloj
  ProcesarConciliacion.Show
End Sub

Private Sub ProcesosContable_Click()
  If ClaveGerente Then
     Control_Procesos Normal, "Fecha de Proceso"
     Cierre.Show
  End If
End Sub

Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
   'Form1.Show
   'Unload FHola
End Sub

Private Sub RecibirXML_Click()
  RatonReloj
  FXMLRecibidosSRI.Show
End Sub

'''Private Sub RPeriodo_Click()
'''   AbrirPeriodo.Show
'''End Sub

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

