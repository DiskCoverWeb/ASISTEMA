VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.MDIForm MDIFact 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Facturacion"
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11400
   Icon            =   "MDIFact.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFact.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dir_Dialog 
      Left            =   420
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   0
      Top             =   105
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7530
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIFact.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIFact.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIFact.frx":2876C
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFact.frx":29046
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
      TabIndex        =   2
      Top             =   7410
      Width           =   11400
      _ExtentX        =   20108
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
         Begin VB.Menu MCAmbioClave 
            Caption         =   "Cambio de Clave"
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
         Caption         =   "Asiento Afiliados"
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
      Begin VB.Menu CompraVentaDivisas 
         Caption         =   "Compra / Venta de Divisas"
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
         Begin VB.Menu MFactReembolso 
            Caption         =   "Facturación de Reembolo de Gastos"
            Shortcut        =   ^R
         End
         Begin VB.Menu MFactReservas 
            Caption         =   "Facturación por Reservas"
         End
         Begin VB.Menu MFacFarmacia 
            Caption         =   "Facturacion de Farmacias"
         End
         Begin VB.Menu MFactDespensa 
            Caption         =   "Despensas"
         End
      End
      Begin VB.Menu MBarrNV 
         Caption         =   "Ingreso de Notas de Venta"
         Begin VB.Menu MNotaVenta 
            Caption         =   "&Nota de Venta"
         End
         Begin VB.Menu MNotaVentaPenc 
            Caption         =   "Nota de Venta de Pensiones"
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
         Begin VB.Menu MLiquidacionCompras 
            Caption         =   "Punto de Venta (Liquidacion de Compras)"
         End
         Begin VB.Menu MPVDonaciones 
            Caption         =   "Punto de Venta (Donaciones)"
         End
         Begin VB.Menu MPuntoVentaTicket 
            Caption         =   "Punto de Venta (Ticket)"
         End
      End
      Begin VB.Menu MAbonoNV 
         Caption         =   "Facturación/Cobros Automática por"
      End
      Begin VB.Menu MRecaudacionAut 
         Caption         =   "Envio/Recepcion Recaudación del Banco"
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
         Shortcut        =   ^D
      End
      Begin VB.Menu SalirSyst 
         Caption         =   "-"
      End
      Begin VB.Menu MImpoExecel 
         Caption         =   "Importaciones desde Excel"
         Shortcut        =   ^E
      End
      Begin VB.Menu MBar5 
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
      Caption         =   "&Herramientas"
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
      Begin VB.Menu EnvioEmail 
         Caption         =   "Prueba de Envio Correo Electronico"
         Shortcut        =   ^W
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
Attribute VB_Name = "MDIFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Clientes_Click()
  RatonReloj
  CliFact = True
  Control_Procesos Normal, "Ingreso a registro de Clientes"
  FClientes.Show
End Sub

'''Private Sub CobDia_Click()
'''  RatonReloj
'''  NuevoComp = True
'''  ModificarComp = False
'''  CopiarComp = False
'''  Co.CodigoB = ""
'''  Co.Numero = 0
'''  FCierreCaja.Show
'''End Sub

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

Private Sub EnvioEmail_Click()
''' Cadena = " HOLA AMIGOS MIOS, ESTA ES UNA PRUENA DE CADENA "
''' MsgBox "         1         2         3         4         5" & vbCrLf _
'''      & "12345678901234567890123456789012345678901234567890" & vbCrLf _
'''      & Cadena & vbCrLf _
'''      & "Left: '" & LeftStrg(Cadena, 10) & "'" & vbCrLf _
'''      & "Right: '" & RightStrg(Cadena, 10) & "'" & vbCrLf _
'''      & "Mid 10,10: '" & MidStrg(Cadena, 10, 10) & "'" & vbCrLf _
'''      & "InStr 'MIOS': '" & InStr(Cadena, "MIOS") & "'"
'''
'''  FEnviarMails.Show
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
    TMail.de = "informacion@diskcoversystem.com"
    TMail.ListaMail = 255
    TMail.TipoDeEnvio = "NN"
    TMail.Asunto = "Prueba de Mails por imap.diskcoversystem.com desde " & Modulo & " [" & Right(CodigoUsuario, 5) & "], Hora (" & Time & ")"
    TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\JAVASCRIPT\f_cartera.html")
    
    html_Informacion_adicional = "<strong>INFORMACION ADICIONAL:</strong><br><br>" _
                               & "<strong>Importe total: USD </strong>150,00<br>" _
                               & "<strong>Importe total: USD </strong>150,00<br>" _
                               & "<strong>Importe total: USD </strong>150,00<br>"
                               
    html_Detalle_adicional = "<tr>" _
                           & "<td>13/12/2024</td>" _
                           & "<td>Prueba de Envio</td>" _
                           & "<td class='row text-right'>150,00</td>" _
                           & "</tr>" _
                           & "<tr>" _
                           & "<td>13/12/2024</td>" _
                           & "<td>Por IMAP</td>" _
                           & "<td class='row text-right'>180,00</td>" _
                           & "</tr>"
    FA.Fecha = FechaSistema
    FA.Recibo_No = Format(FA.Fecha, "yyyymmdd") & Format(FA.Factura, "000000000")
    TMail.Adjunto = "C:\SYSBASES\TEMP\archivo.xml"
    TMail.para = ""
    Insertar_Mail TMail.para, "actualizar@diskcoversystem.com"
    Insertar_Mail TMail.para, "diskcoversystem@msn.com"
    Insertar_Mail TMail.para, "diskcover.system@yahoo.com"
    Insertar_Mail TMail.para, "diskcover.system@gmail.com"
    FEnviarCorreos.Show 1
    MsgBox "Correos Enviados a (" & Si_No & "):" & vbCrLf & Replace(TMail.para, ";", ";" & vbCrLf) & vbCrLf _
         & String(70, "_") & vbCrLf & vbCrLf _
         & "De: " & TMail.de & vbCrLf & vbCrLf _
         & "Asunto: " & TMail.Asunto & vbCrLf
    TMail.para = ""
    TMail.ListaMail = 0
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

Private Sub MAbonoNV_Click()
  RatonReloj
  FRecaudacionBancosCxC.Show
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
  Control_Procesos Normal, "Asiento de Afiliados"
  FClientesRazonSocial.Show
End Sub

Private Sub MCambCod_Click()
  If ClaveSupervisor Then
     RatonReloj
     Opciones = 1
     FCambiosCodigos.Show
  End If
End Sub

Private Sub MCAmbioClave_Click()
    Titulo = "CAMBIO DE CLAVE"
    Mensajes = "Estimado " & NombreUsuario & ", desea cambiar su clave de acceso?"
    If BoxMensaje = vbYes Then
       RatonReloj
       Control_Procesos Normal, "Cambio de Clave"
       CambClav.Show
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

Private Sub MCierresDeCaja_Click()
  RatonReloj
  NuevoComp = True
  ModificarComp = False
  CopiarComp = False
  Co.CodigoB = ""
  Co.Numero = 0
  FCierreCaja.Show
End Sub

Private Sub MCobProg_Click()
  FCobrosProgramados.Show
End Sub

Private Sub MCtasDocxCobrar_Click()
  RatonReloj
  CxCNivel.Show
End Sub

Private Sub MFacFarmacia_Click()
  RatonReloj
  TipoFactura = "FA"
  FacturaNueva = True
  FacturaFarmacia.Show
End Sub

Private Sub MFactDespensa_Click()
  RatonReloj
  TipoFactura = "DES"
  FacturaNueva = True
  FacturaDespensa.Show
End Sub

Private Sub MFactReembolso_Click()
  RatonReloj
  TipoFactura = "FR"
  FacturaNueva = True
  FacturaReembolso.Show
End Sub

Private Sub MRecaudacionAut_Click()
  RatonReloj
  FRecaudacionBancosPreFa.Show
End Sub

Private Sub Timer1_Timer()
    If TiempoTarea = 0 Then TiempoTarea = Time
    If TiempoSistema = 0 Then TiempoSistema = Time
    If TiempoServidor = 0 Then TiempoServidor = Time
    MiTiempo = Time
    MiTiempo = CSng(Format$(Minute(MiTiempo - TiempoServidor), "00") & "." & Format$(Second(MiTiempo - TiempoServidor), "00"))
'''    If MiTiempo >= 0.59 Then
'''       Select Case ContadorServidor
'''         Case 1: 'Conectamos el Socket: ServidorMySQL
'''                 MDIWinsock.Close
'''                 MDIWinsock.Connect "mysql.diskcoversystem.com", 13306
'''         Case 2: 'Conectamos el Socket: ServidorSQLServer
'''                 MDIWinsock.Close
'''                 MDIWinsock.Connect "mysql.diskcoversystem.com", 11433
'''         Case 3: 'Conectamos el Socket: ServidorSRIPrueba
'''                 MDIWinsock.Close
'''                 MDIWinsock.Connect "celcer.sri.gob.ec", 443
'''
'''         Case 4: 'Conectamos el Socket: ServidorSRIProduccion
'''                 MDIWinsock.Close
'''                 MDIWinsock.Connect "cel.sri.gob.ec", 443
'''       End Select
''''''        MsgBox ContadorServidor & vbCrLf _
''''''               & ServidorMySQL & vbCrLf _
''''''               & ServidorSQLServer & vbCrLf _
''''''               & ServidorSRIPrueba & vbCrLf _
''''''               & ServidorSRIProduccion
'''
'''       TiempoServidor = Time
'''       ContadorServidor = ContadorServidor + 1
'''       If ContadorServidor > 4 Then ContadorServidor = 1
'''    End If
    Ver_Grafico_FormPict
    If CodigoUsuario <> Ninguno Then Recordar_Tarea_Hora
End Sub

Private Sub MDIForm_Activate()
    screen_size
End Sub

Private Sub MDIForm_Load()
  TiempoTarea = Time
  TiempoSistema = Time
  TiempoServidor = Time
  Timer1.Interval = 1000
  
  ContadorServidor = 1
  ServidorMySQL = False
  ServidorSQLServer = False
  ServidorSRIPrueba = False
  ServidorSRIProduccion = False

  Set MDIFormulario = Me
  Primera_Vez = True
  Bandera = True
  UnidadSistema
 ' TipoModulo = Factu
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "FACTURACION"
  MenuDeModulos = True
  IngresarClave = True
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Facturacion"
  End
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

Private Sub MImpoExecel_Click()
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

Private Sub MLiquidacionCompras_Click()
  RatonReloj
  TipoFactura = "LC"
  FacturaNueva = True
  FacturasPV.Show
End Sub

'''Private Sub MListAlumSistemaEdu_Click()
'''  RatonReloj
'''  FListFox.Show
'''End Sub

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
  FNotasDeCredito.Show
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

Private Sub MPVDonaciones_Click()
  RatonReloj
  TipoFactura = "DO"
  FacturaNueva = True
  FacturasPV.Show
End Sub

Private Sub MReciboAbonoAnticipado_Click()
  MayorAux.Show
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

Private Sub MTraslVentAnt_Click()
  RatonReloj
  AsientoAuto.Show
End Sub

Private Sub MWebServices_Click()
   Form_WS.Show
End Sub

Private Sub Mwwwdiskcover_Click()
Dim iRet As Long
  Control_Procesos "Q", "Salir por ingresar a la Pagina WEB"
  MsgBox "Estas a punto de ingresar al Centro de descargas del sistema"
  iRet = Shell("rundll32.exe url.dll,FileProtocolHandler " & "https://www.diskcoversystem.com", vbMaximizedFocus)
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

