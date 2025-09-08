VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.MDIForm MDIFarmacia 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Facturacion"
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11400
   Icon            =   "MDIFarmacia.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFarmacia.frx":164A
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
            Picture         =   "MDIFarmacia.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIFarmacia.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIFarmacia.frx":2876C
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFarmacia.frx":29046
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
   Begin MSComDlg.CommonDialog Dir_Dialog 
      Left            =   525
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu BaseRel 
      Caption         =   "&Archivos"
      Begin VB.Menu MDelSistema 
         Caption         =   "Del Sistema"
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
         Begin VB.Menu MImpFarmacia 
            Caption         =   "Importar Clientes Farmacia"
         End
         Begin VB.Menu MCAmbioClave 
            Caption         =   "Cambio de Clave"
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
      End
      Begin VB.Menu MCtasDocxCobrar 
         Caption         =   "Ctas. y Dtos. por Cobrar (CLIENTES)"
         Visible         =   0   'False
      End
      Begin VB.Menu Mbarr1 
         Caption         =   "-"
      End
      Begin VB.Menu MFacFarmacia 
         Caption         =   "Facturacion de Farmacias"
         Shortcut        =   ^F
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
         Caption         =   "Importar Abonos por Excel en bloque"
         Shortcut        =   ^I
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
      Begin VB.Menu MCatProd 
         Caption         =   "Catalogo de Productos"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MImpRecCaja 
         Caption         =   "Imprimir Recibo de Caja"
      End
      Begin VB.Menu MBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MHistFact 
         Caption         =   "Resumen de Cartera/Historico de Facturas/Ventas"
         Shortcut        =   ^H
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
Attribute VB_Name = "MDIFarmacia"
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

Private Sub ListarFact_Click()
   RatonReloj
   ListFact.Show
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

Private Sub MAsientoMatricula_Click()
  RatonReloj
  CliFact = True
  Control_Procesos Normal, "Asiento de Afiliados"
  FClientesRazonSocial.Show
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

Private Sub MCambioPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     CambioPeriodo.Show 1
     PonerDirEmpresa
  End If
End Sub

Private Sub MCatProd_Click()
  RatonReloj
  CatalogoCtas.Show
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

Private Sub MImpFarmacia_Click()
  RatonReloj
  FImpFarm.Show
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
  Modulo = "FARMACIA"
  MenuDeModulos = True
  IngresarClave = True
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Facturacion"
  End
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

Private Sub MIngLinea_Click()
  RatonReloj
  IngLinea.Show
End Sub

Private Sub MNotasCreditos_Click()
  FNotasDeCredito.Show
End Sub

Private Sub MRetencion_Click()
  RatonReloj
  AbonoRetencion.Show
End Sub

Private Sub Mwwwdiskcover_Click()
Dim iRet As Long
  Control_Procesos "Q", "Salir por ingresar a la Pagina WEB"
  MsgBox "Estas a punto de ingresar al Centro de descargas del sistema"
  iRet = Shell("rundll32.exe url.dll,FileProtocolHandler " & "https://www.diskcoversystem.com", vbMaximizedFocus)
  End
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

Private Sub SalirSystem_Click()
  Control_Procesos "Q", "Salir Modulo de Facturacion"
  End
End Sub

