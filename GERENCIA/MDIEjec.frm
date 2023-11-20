VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIEjec 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Facturacion"
   ClientHeight    =   6690
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11025
   Icon            =   "MDIEjec.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIEjec.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11025
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11025
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
      Top             =   6315
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIEjec.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIEjec.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIEjec.frx":2876C
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIEjec.frx":28A86
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIEjec.frx":29360
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
      Top             =   6195
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu BaseRel 
      Caption         =   "&Archivos"
      Begin VB.Menu MDelSistema 
         Caption         =   "Del Sistema"
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
         Begin VB.Menu MModAud 
            Caption         =   "&Modulos de Auditoria"
         End
         Begin VB.Menu MCambCod 
            Caption         =   "Cambio de Codigos de Facturacion"
         End
         Begin VB.Menu MCambioComision 
            Caption         =   "Cambio de Codigos de Comisiones"
         End
         Begin VB.Menu MCamboNumFact 
            Caption         =   "Cambio de Numero en la Factura"
         End
      End
      Begin VB.Menu Clientes 
         Caption         =   "&Clientes"
         Shortcut        =   ^B
      End
      Begin VB.Menu MAsientoMatricula 
         Caption         =   "Datos del Representante"
         Shortcut        =   ^A
      End
      Begin VB.Menu MSetCtas 
         Caption         =   "Seteos de Cuentas y Articulos"
         Begin VB.Menu MIngLinea 
            Caption         =   "&CxC Clientes/NC/Autorizaciones"
         End
         Begin VB.Menu Productos 
            Caption         =   "Ingresar &Productos (Ventas)"
         End
         Begin VB.Menu MActualizarVentas 
            Caption         =   "Actualizar Ventas Anticipadas"
         End
      End
      Begin VB.Menu MCtasDocxCobrar 
         Caption         =   "Ctas. y Dtos. por Cobrar (CLIENTES)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mbarr1 
         Caption         =   "-"
      End
      Begin VB.Menu MOrdenProduccion 
         Caption         =   "Ingresar Orden de Produccion"
      End
      Begin VB.Menu MOpcDiv 
         Caption         =   "-"
      End
      Begin VB.Menu CompraVentaDivisas 
         Caption         =   "Compra / Venta de Divisas"
      End
      Begin VB.Menu MBarAutos 
         Caption         =   "-"
      End
      Begin VB.Menu MFleCredPedOtros 
         Caption         =   "Fletes/Creditos/Pedidos/Otros"
         Begin VB.Menu MIngFletes 
            Caption         =   "Ingreso de Fletes"
         End
         Begin VB.Menu MPedidos 
            Caption         =   "Pedidos"
            Begin VB.Menu MHabitaciones 
               Caption         =   "Ingresar Pedidos"
            End
            Begin VB.Menu MCambioPedidos 
               Caption         =   "Cambios de Pedidos"
            End
         End
         Begin VB.Menu MIngCredito 
            Caption         =   "Ingreso de Crédito"
         End
      End
      Begin VB.Menu MBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MPuntoVentas 
         Caption         =   "Ingreso de Facturas"
         Begin VB.Menu IngFact 
            Caption         =   "&Facturacion"
            Shortcut        =   ^F
         End
         Begin VB.Menu MFacturarPensiones 
            Caption         =   "Facturacion de Pensiones"
            Shortcut        =   ^P
         End
         Begin VB.Menu MFacturarMatriculas 
            Caption         =   "Facturacion de Matriculas"
            Shortcut        =   ^M
         End
         Begin VB.Menu MFactServ 
            Caption         =   "Facturación de Servicios"
            Shortcut        =   ^Y
         End
         Begin VB.Menu MFactReservas 
            Caption         =   "Facturación por Reservas"
         End
      End
      Begin VB.Menu MBarrNV 
         Caption         =   "Ingreso de Notas de Venta"
         Begin VB.Menu MNotaVenta 
            Caption         =   "&Nota de Venta"
         End
         Begin VB.Menu MNotaVentaPenc 
            Caption         =   "Nota de Venta de Pensiones"
            Shortcut        =   ^R
         End
         Begin VB.Menu MIngFactDUI 
            Caption         =   "Ingresar Facturas DUI"
         End
      End
      Begin VB.Menu MBar4 
         Caption         =   "Ingreso de Punto de Ventas"
         Begin VB.Menu MIngFactHotel 
            Caption         =   "Punto de Venta (Facturas)"
         End
         Begin VB.Menu MPuntoVenta 
            Caption         =   "Punto de Venta (Nota de Venta)"
         End
         Begin VB.Menu MPuntoVentaTicket 
            Caption         =   "Punto de Venta (Ticket)"
         End
         Begin VB.Menu MLiquidacionCompras 
            Caption         =   "Punto de Venta (Liquidacion de Compras)"
         End
      End
      Begin VB.Menu MAbonoNV 
         Caption         =   "Facturación/Cobros Automática por"
         Begin VB.Menu MRespRest 
            Caption         =   "&Banco Bolivariano"
         End
         Begin VB.Menu MBcoInter 
            Caption         =   "Banco &Internacional"
         End
         Begin VB.Menu MBcoPichinha 
            Caption         =   "Banco del &Pichincha"
         End
         Begin VB.Menu MBcoPacific 
            Caption         =   "Banco del &Pacífico"
         End
         Begin VB.Menu MIntermatico 
            Caption         =   "Banco del Pacifico (Inter&matico)"
         End
         Begin VB.Menu MProdubanco 
            Caption         =   "Banco Produbanco"
         End
         Begin VB.Menu MBGR_EC 
            Caption         =   "Banco General Rumiñahui"
         End
         Begin VB.Menu MBancoGuayaquil 
            Caption         =   "Banco de Guayaquil"
         End
         Begin VB.Menu MBar6 
            Caption         =   "-"
         End
         Begin VB.Menu MAbonoAutomatocs 
            Caption         =   "Abonos Automaticos"
         End
         Begin VB.Menu MCooperativas 
            Caption         =   "-"
         End
         Begin VB.Menu MCoopJep 
            Caption         =   "Cooperativa JEP"
         End
         Begin VB.Menu MCoopCACPE 
            Caption         =   "Cooperativa CACPE Bliblian Ltda"
         End
      End
      Begin VB.Menu MRecaudacionAut 
         Caption         =   "Recaudación de Facturas Automátizada por"
         Begin VB.Menu MBcoRecBolivariano 
            Caption         =   "Recaudacion Banco Bolivariano"
         End
         Begin VB.Menu MBcoRecInternacional 
            Caption         =   "Recaudacion Banco Internacional"
         End
         Begin VB.Menu MBcoGuayquil 
            Caption         =   "Recaudacion Banco de Guayaquil"
         End
         Begin VB.Menu BcoRecPichincha 
            Caption         =   "Recaudacion Banco del Pichincha"
         End
         Begin VB.Menu MBcoPacificoBizbanck 
            Caption         =   "Recaudacion Banco del Pacifico"
         End
         Begin VB.Menu MRecProdubanco 
            Caption         =   "Recaudacion del Produbanco"
         End
         Begin VB.Menu MRecOtrosBancos 
            Caption         =   "Recaudacion de Otros Bancos"
         End
         Begin VB.Menu MRecTarjetas 
            Caption         =   "Recaudacion de Tarjetas"
         End
         Begin VB.Menu MBCoopJep1 
            Caption         =   "-"
         End
         Begin VB.Menu MCoopJep1 
            Caption         =   "Cooperativa JEP"
         End
         Begin VB.Menu MSubirFarmacia 
            Caption         =   "Subir Farmacias"
         End
         Begin VB.Menu MRecaudacionxExcel 
            Caption         =   "Recaudacion por Excel"
         End
      End
      Begin VB.Menu MBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MImpoExecel 
         Caption         =   "Importaciones desde Excel"
         Begin VB.Menu MImportarVentasExcel 
            Caption         =   "Importar Ventas desde Excel"
         End
         Begin VB.Menu MImportAbonosBloc 
            Caption         =   "Importar Abonos en bloque"
         End
         Begin VB.Menu MFactConsumo 
            Caption         =   "Facturacion de Consumos"
         End
      End
      Begin VB.Menu MBAbono1 
         Caption         =   "-"
      End
      Begin VB.Menu MAbonosAntiCli 
         Caption         =   "Abonos Anticipados de Clientes (CxP)"
      End
      Begin VB.Menu MChequesProtestados 
         Caption         =   "Cheques Protestados/Devueltos"
      End
      Begin VB.Menu MNotasCreditos 
         Caption         =   "Notas de Credito"
         Shortcut        =   ^N
      End
      Begin VB.Menu MAbonosFact 
         Caption         =   "Abonos de Facturas/Notas de Venta"
         Begin VB.Menu CtasCobrar 
            Caption         =   "Detallado de CxC/Retenciones"
            Shortcut        =   {F5}
         End
         Begin VB.Menu MRetencion 
            Caption         =   "Retenciones"
            Shortcut        =   {F6}
         End
         Begin VB.Menu MAbonoEfect 
            Caption         =   "En Efectivo"
            Shortcut        =   {F7}
         End
         Begin VB.Menu MAbonoDep 
            Caption         =   "Por Depósito"
            Shortcut        =   {F8}
         End
         Begin VB.Menu MPorTransf 
            Caption         =   "Por Transferencias"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu MAnticipos 
            Caption         =   "Por Anticipos"
            Shortcut        =   {F9}
         End
         Begin VB.Menu PorAnticiposCxC 
            Caption         =   "Por Anticipos en CxC"
         End
         Begin VB.Menu MPorAntGrupo 
            Caption         =   "Por Anticipos en Grupo"
         End
         Begin VB.Menu MAbonoCredito 
            Caption         =   "Por Tarjeta de Créditos"
            Shortcut        =   {F11}
         End
         Begin VB.Menu MPorDiferencias 
            Caption         =   "Por Diferencias"
         End
         Begin VB.Menu Mx1 
            Caption         =   "-"
         End
         Begin VB.Menu MAbonoAuto 
            Caption         =   "Abonos Automaticos en Grupo"
         End
         Begin VB.Menu MLotesBauchers 
            Caption         =   "Cancelacion de Lotes/Bauchers"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu ListarFact 
         Caption         =   "Listar o Anular: Facturas/Nota de Ventas/Ordenes"
         Shortcut        =   ^L
      End
      Begin VB.Menu MBAbono2 
         Caption         =   "-"
      End
      Begin VB.Menu MCierresDeCaja 
         Caption         =   "Cierre de Caja"
         Begin VB.Menu MCierreCajaTicket 
            Caption         =   "Por Tickete"
            Shortcut        =   ^K
         End
         Begin VB.Menu CobDia 
            Caption         =   "Cierre Diario"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu SalirSyst 
         Caption         =   "-"
      End
      Begin VB.Menu ProcOtraEmp 
         Caption         =   "Cambiar de &Empresa"
      End
      Begin VB.Menu SalirSystem 
         Caption         =   "&Salir del Sistema"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MListGrupos 
         Caption         =   "Listar Clientes por Grupo"
         Shortcut        =   ^G
      End
      Begin VB.Menu LisTCheqPosf 
         Caption         =   "Listar Cheques posfechados"
         Visible         =   0   'False
      End
      Begin VB.Menu MCatProd 
         Caption         =   "Catalogo de Productos"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MImpRecCaja 
         Caption         =   "Imprimir Recibo de Caja"
      End
      Begin VB.Menu MReciboAbonoAnticipado 
         Caption         =   "Imprimir Recibo de Abonos Anticipados"
      End
      Begin VB.Menu MBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MHistFact 
         Caption         =   "Resumen de Cartera/Historico de Facturas/Ventas"
         Shortcut        =   ^H
      End
      Begin VB.Menu MResumenMensual 
         Caption         =   "Resumen de Produccion Mensuales"
      End
   End
   Begin VB.Menu MOtrosProc 
      Caption         =   "&Otros Procesos"
      Begin VB.Menu MCobProg 
         Caption         =   "Cobros programados"
      End
      Begin VB.Menu ResumCratT 
         Caption         =   "Suscripciones por Periodos"
      End
      Begin VB.Menu MResumSuscrip 
         Caption         =   "Resumen de Suscripciones"
      End
      Begin VB.Menu MNominaAlumnos 
         Caption         =   "Nomina de Alumnos"
      End
      Begin VB.Menu MSaldoAlumnoss 
         Caption         =   "Estado de Cartera de Alumnos"
      End
      Begin VB.Menu MResumComis 
         Caption         =   "Comisiones"
      End
      Begin VB.Menu MResumFletes 
         Caption         =   "Resumen de Fletes"
      End
      Begin VB.Menu MBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MTraslVentAnt 
         Caption         =   "Traspaso Ventas Anticipadas a Ventas"
      End
      Begin VB.Menu BarLstAlumSistEdu 
         Caption         =   "-"
      End
      Begin VB.Menu MListAlumSistemaEdu 
         Caption         =   "Listado Alumnos Sistema Educativo"
      End
      Begin VB.Menu MGenerarPDF 
         Caption         =   "General PDF"
      End
      Begin VB.Menu MWebServices 
         Caption         =   "Web Services"
      End
   End
   Begin VB.Menu Mwwwdiskcover 
      Caption         =   "Visitanos en: www.diskcoversystem.com"
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "Ambiente"
      Enabled         =   0   'False
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "MDIEjec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BcoRecPichincha_Click()
  RatonReloj
  TextoBanco = "PICHINCHA"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub Clientes_Click()
  RatonReloj
  CliFact = True
  Control_Procesos Normal, "Ingreso a registro de Clientes"
  FClientes.Show
End Sub

Private Sub CobDia_Click()
  RatonReloj
  NuevoComp = True
  ModificarComp = False
  CopiarComp = False
  Co.CodigoB = ""
  Co.Numero = 0
  FCierreCaja.Show
End Sub

Private Sub CompraVentaDivisas_Click()
  RatonReloj
  FacturaNueva = True
  FacturasDiv.Show
End Sub

Private Sub CtasCobrar_Click()
   RatonReloj
   With FA
       .TC = Ninguno
       .Autorizacion = Ninguno
       .Serie = Ninguno
       .CodigoC = Ninguno
       .Factura = 0
   End With
   Abonos.Show

End Sub

Private Sub IngFact_Click()
  RatonReloj
  Control_Procesos Pendiente, "Procesar Factura"
  OpcServicio = False
  TipoFactura = "FA"
  FacturaNueva = True
  Facturas.Show
End Sub

Private Sub ListarFact_Click()
   RatonReloj
   ListFact.Show
End Sub

Private Sub LisTCheqPosf_Click()
  RatonReloj
  PosFecha.Show
End Sub

Private Sub MAbonoAuto_Click()
  RatonReloj
  TipoDoc = ""
  AbonoAutomatico.Show
End Sub

Private Sub MAbonoAutomatocs_Click()
  RatonReloj
  TextoBanco = "AUTOMATICOS"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MAbonoCredito_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "TARJETA"
  AbonoAnticipo.Show 1
End Sub

Private Sub MAbonoDep_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "BANCOS"
  AbonoAnticipo.Show 1
End Sub

Private Sub MAbonoEfect_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "EFECTIVO"
  AbonoAnticipo.Show 1
End Sub

Private Sub MAbonosAntiCli_Click()
   TipoFactura = "AA"
   AbonoAnticipado.Show 1
End Sub

Private Sub MActualizarVentas_Click()
  RatonReloj
  AcEntSal.Show
End Sub

Private Sub MAnticipos_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "ANTICIPOS"
  AbonoAnticipo.Show 1
End Sub

Private Sub MAsientoMatricula_Click()
  RatonReloj
  CliFact = True
  Control_Procesos Normal, "Asiento de Matriculas"
  FClientesRazonSocial.Show
End Sub

Private Sub MBancoGuayaquil_Click()
  RatonReloj
 'FBancoPacifico.Show
  TextoBanco = "GUAYAQUIL"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MBcoGuayquil_Click()
  RatonReloj
  TextoBanco = "GUAYAQUIL"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MBcoInter_Click()
  RatonReloj
 'FBancoInternacional.Show
  TextoBanco = "INTERNACIONAL"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MBcoPacific_Click()
  RatonReloj
 'FBancoPacifico.Show
  TextoBanco = "PACIFICO"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MBcoPacificoBizbanck_Click()
  RatonReloj
 'FBancoPichincha.Show
  TextoBanco = "BIZBANCKPACIFICO"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MBcoPichinha_Click()
  RatonReloj
  TextoBanco = "PICHINCHA"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MBcoRecBolivariano_Click()
  RatonReloj
  TextoBanco = "BOLIVARIANO"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MBcoRecInternacional_Click()
  RatonReloj
  TextoBanco = "INTERNACIONAL"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MBGR_EC_Click()
  RatonReloj
  'FBancoPichincha.Show
  TextoBanco = "BGR_EC"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MCambCod_Click()
  If ClaveSupervisor Then
     RatonReloj
     Opciones = 1
     FCambiosCodigos.Show
  End If
End Sub

Private Sub MCambioComision_Click()
  If ClaveSupervisor Then
     RatonReloj
     Opciones = 2
     FCambiosCodigos.Show
  End If
End Sub

Private Sub MCambioPedidos_Click()
  If ClaveSupervisor Then
     RatonReloj
     FCambioPedidos.Show
  End If
End Sub

Private Sub MCambioPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
End Sub

Private Sub MCamboNumFact_Click()
  If ClaveSupervisor Then
     RatonReloj
     Opciones = 3
     FCambiosCodigos.Show
  End If
End Sub

Private Sub MCatProd_Click()
  RatonReloj
  CatalogoCtas.Show
End Sub

Private Sub MChequesProtestados_Click()
  RatonReloj
  TipoFactura = "CP"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MCierreCajaTicket_Click()
  RatonReloj
  FConvertirPV.Show
End Sub

Private Sub MCobProg_Click()
  FCobrosProgramados.Show
End Sub

Private Sub MCoopCACPE_Click()
  RatonReloj
  TextoBanco = "CACPE"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MCoopJep_Click()
  RatonReloj
  TextoBanco = "COOPJEP"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MCoopJep1_Click()
  RatonReloj
  TextoBanco = "COOPJEP"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MCtasDocxCobrar_Click()
  RatonReloj
  CxCNivel.Show
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
 ' TipoModulo = Factu
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "EJECUTIVOS"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 1000
  IngresarClave = True
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Facturacion"
  End
End Sub

Private Sub MFactConsumo_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub MFactServ_Click()
  RatonReloj
  Control_Procesos Pendiente, "Procesar Factura"
  OpcServicio = True
  TipoFactura = "PV"
  FacturaNueva = True
  Facturas.Show
End Sub

Private Sub MFacturarMatriculas_Click()
  RatonReloj
  Control_Procesos Pendiente, "Procesar Factura de Matricula"
  OpcServicio = False
  FacturaMatricula = True
  TipoFactura = "FA"
  FacturaNueva = True
  FacturasPension.Show
End Sub

Private Sub MFacturarPensiones_Click()
  RatonReloj
  Control_Procesos Pendiente, "Procesar Factura de Pension"
  OpcServicio = False
  FacturaMatricula = False
  TipoFactura = "FA"
  FacturaNueva = True
  FacturasPension.Show
End Sub

Private Sub MGenerarPDF_Click()
   FGeneraPDF.Show
End Sub

Private Sub MHabitaciones_Click()
  RatonReloj
  'FHabitacion.Show
  FPedidos.Show
End Sub

Private Sub MHistFact_Click()
  RatonReloj
  Control_Procesos Normal, "Historial de Facturas"
  HistorialFacturas.Show
End Sub

Private Sub MImportAbonosBloc_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub MImportarVentasExcel_Click()
   RatonReloj
   FImporta.Show
End Sub

Private Sub MImpRecCaja_Click()
  RatonReloj
  RAbonos.Show
End Sub

Private Sub MIngCredito_Click()
  RatonReloj
  FacturaCredito.Show
End Sub

Private Sub MIngFactDUI_Click()
  RatonReloj
  OpcServicio = False
  TipoFactura = "FA"
  FacturasDUI.Show
End Sub

Private Sub MIngFactHotel_Click()
  RatonReloj
  TipoFactura = "FA"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MIngFletes_Click()
  RatonReloj
  FFletes.Show
End Sub

Private Sub MIngLinea_Click()
  RatonReloj
  IngLinea.Show
End Sub

Private Sub MIntermatico_Click()
  RatonReloj
 'FBancoPacifico.Show
  TextoBanco = "INTERMATICO"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MLiquidacionCompras_Click()
  RatonReloj
  TipoFactura = "LC"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MListAlumSistemaEdu_Click()
  RatonReloj
  FListFox.Show
End Sub

Private Sub MListGrupos_Click()
  RatonReloj
  Control_Procesos Normal, "Proceso de Facturacion Multiple"
  ListarGrupos.Show
End Sub

Private Sub MLotesBauchers_Click()
  FPagoTJ.Show
End Sub

Private Sub MModAud_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Ingreso al Modulo de Auditoria"
     FAuditoria.Show
  End If
End Sub

Private Sub MNominaAlumnos_Click()
  RatonReloj
  FMatriculados.Show
End Sub

Private Sub MNotasCreditos_Click()
  FNC.Show
End Sub

Private Sub MNotaVenta_Click()
  RatonReloj
  OpcServicio = False
  TipoFactura = "NV"
  FacturaNueva = True
  Facturas.Show
End Sub

Private Sub MNotaVentaPenc_Click()
  RatonReloj
  OpcServicio = False
  TipoFactura = "NV"
  FacturaNueva = True
  FacturasPension.Show
End Sub

Private Sub MOrdenProduccion_Click()
  RatonReloj
  Control_Procesos Pendiente, "Procesar Orden de Produccion"
  OpcServicio = False
  TipoFactura = "OP"
  FacturaNueva = True
  Facturas.Show
End Sub

Private Sub MPorAntGrupo_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "ANTICIPOS"
  AbonoAnticipoGrupo.Show
End Sub

Private Sub MPorDiferencias_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "DIFERENCIAS"
  AbonoAnticipo.Show 1
End Sub

Private Sub MPorTransf_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "TRANSFERENCIA"
  AbonoAnticipo.Show 1
End Sub

Private Sub MProdubanco_Click()
  RatonReloj
  TextoBanco = "PRODUBANCO"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MPuntoVenta_Click()
  RatonReloj
  TipoFactura = "NV"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MPuntoVentaTicket_Click()
  RatonReloj
  TipoFactura = "PV"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MRecaudacionxExcel_Click()
  RatonReloj
  TextoBanco = "POREXCEL"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MReciboAbonoAnticipado_Click()
  MayorAux.Show
End Sub

Private Sub MRecOtrosBancos_Click()
  RatonReloj
  TextoBanco = "OTROSBANCOS"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MRecProdubanco_Click()
  RatonReloj
  TextoBanco = "PRODUBANCO"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MRecTarjetas_Click()
  RatonReloj
  TextoBanco = "TARJETAS"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MRespRest_Click()
  RatonReloj
  TextoBanco = "BOLIVARIANO"
  FRecaudacionBancosCxC.Show
End Sub

Private Sub MResumComis_Click()
  RatonReloj
  FComisiones.Show
End Sub

Private Sub MResumenMensual_Click()
  RatonReloj
  ResumenProduccion.Show
End Sub

Private Sub MResumFletes_Click()
  RatonReloj
' FResumenFletes.Show
  FGeneraPDF.Show
End Sub

Private Sub MResumSuscrip_Click()
  RatonReloj
  FResSusc.Show
End Sub

Private Sub MRetencion_Click()
  RatonReloj
  AbonoRetencion.Show
End Sub

Private Sub MSaldoAlumnoss_Click()
  RatonReloj
  FSaldosClientes.Show
End Sub

Private Sub MSubirFarmacia_Click()
  RatonReloj
  TextoBanco = "FARMACIAS"
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub MTraslVentAnt_Click()
  RatonReloj
  AsientoAuto.Show
End Sub

Private Sub MWebServices_Click()
   Form_WS.Show
End Sub

Private Sub Mwwwdiskcover_Click()
Dim X
  Control_Procesos "Q", "Salir por ingresar a la Pagina WEB"
  MsgBox "Estas a punto de ingresar al Centro de descargas del sistema"
  X = ShellExecute(Me.hwnd, "Open", "http://www.diskcoversystem.com", &O0, &O0, SW_NORMAL)
  End
End Sub

Private Sub PorAnticiposCxC_Click()
  RatonReloj
  TipoDoc = ""
  Nuevo = True
  TipoProc = "ANTICIPOSCXC"
  AbonoAnticipo.Show 1
End Sub

Private Sub ProcOtraEmp_Click()
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub Productos_Click()
   RatonReloj
   IngProdInv.Show
End Sub

Private Sub ResumCratT_Click()
   RatonReloj
   ListarSuscripciones.Show
End Sub

Private Sub SalirSystem_Click()
  Control_Procesos "Q", "Salir Modulo de Facturacion"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  'And CodigoUsuario <> Ninguno
  Recordar_Tarea_Hora
End Sub
