VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.MDIForm MDICosto 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   7935
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11700
   Icon            =   "MDICosto.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDICosto.frx":164A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   11700
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11700
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
      Top             =   7560
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDICosto.frx":27808
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDICosto.frx":27E92
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDICosto.frx":2876C
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDICosto.frx":29046
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
      Top             =   7440
      Width           =   11700
      _ExtentX        =   20638
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
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Manten 
         Caption         =   "Del Sistema"
         Begin VB.Menu MCambioPeriodo 
            Caption         =   "Cambio de Periodo"
         End
         Begin VB.Menu MCambClave 
            Caption         =   "Cambio de Clave"
         End
      End
      Begin VB.Menu Procesos 
         Caption         =   "De Operacion"
         Begin VB.Menu IngArts 
            Caption         =   "Ingreso de Articulos"
            Shortcut        =   ^A
         End
         Begin VB.Menu MIngBodega 
            Caption         =   "Ingreso de Bodegas"
            Shortcut        =   ^B
         End
         Begin VB.Menu MIngMarcas 
            Caption         =   "Ingreso de Marcas"
            Shortcut        =   ^M
         End
         Begin VB.Menu MIngProv 
            Caption         =   "Ingresar Proveedores"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu x1 
            Caption         =   "-"
         End
         Begin VB.Menu ControlInv 
            Caption         =   "Control de Inventario (E/S)"
            Shortcut        =   ^I
         End
         Begin VB.Menu MSalidaCero 
            Caption         =   "Control de Salida de Bodega"
            Shortcut        =   ^S
         End
         Begin VB.Menu MTransf_Bodegas 
            Caption         =   "Transferencia de Productos Bodegas"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu MCambioSalida 
            Caption         =   "Transferencia de Productos Bodegas (Recetas)"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu MMercaderiaPP 
            Caption         =   "Inventarios de Productos en Procesos"
            Shortcut        =   ^P
         End
         Begin VB.Menu MInvProdTerminados 
            Caption         =   "Inventarios de Productos Terminados"
            Shortcut        =   ^R
         End
         Begin VB.Menu MIngCosteo 
            Caption         =   "Ingreso de Mercaderia por Costeo"
         End
         Begin VB.Menu MIngMercDUI 
            Caption         =   "Ingreso/Egresos de Mercadería por lotes"
         End
      End
      Begin VB.Menu x 
         Caption         =   "-"
      End
      Begin VB.Menu MImpoDifExcel 
         Caption         =   "Importar Diferencias desde Excell"
      End
      Begin VB.Menu MB1 
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
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu MMayorProd 
         Caption         =   "Mayorizar Productos"
         Shortcut        =   ^T
      End
      Begin VB.Menu ListProd 
         Caption         =   "Listar Productos"
      End
      Begin VB.Menu ExitenInv 
         Caption         =   "Kardex de Productos"
         Shortcut        =   ^K
      End
      Begin VB.Menu salinvart 
         Caption         =   "Notas de Entrada/Salida"
      End
      Begin VB.Menu MSalidaPVP 
         Caption         =   "Costos con PVP"
      End
      Begin VB.Menu ResumExist 
         Caption         =   "Resumen de Existencia"
         Shortcut        =   ^E
      End
      Begin VB.Menu MResumenTotales 
         Caption         =   "Resumen de Compras/Ventas Promediadas"
      End
   End
   Begin VB.Menu Herramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MReindexar_Kardex 
         Caption         =   "Reindexar Kardex"
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "Ambiente"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDICosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub CambiaEmp_Click()
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub ControlInv_Click()
  RatonReloj
  Empleados = False
  Kard_Ing.Show
End Sub

Private Sub ExitenInv_Click()
  RatonReloj
  Existen.Show
End Sub

Private Sub IngArts_Click()
  RatonReloj
  IngProdInv.Show
End Sub

Private Sub ListProd_Click()
  RatonReloj
  CatalogoCtas.Show
End Sub

Private Sub MCambClave_Click()
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

Private Sub MCambioSalida_Click()
  If ClaveAuxiliar Then
     RatonReloj
     Kard_Cambio.Show
  End If
End Sub

Private Sub MDIForm_Activate()
    screen_size
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
  Primera_Vez = True
  Bandera = True
  UnidadSistema
  IngresarClave = True
  NumModulo = "0"
  Modulo = "INVENTARIO"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 10000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Inventario"
  End
End Sub

Private Sub MImpoDifExcel_Click()
  If ClaveContador Then
     RatonReloj
     Control_Procesos Normal, "Importar diferencias desde Excel"
     FImporta.Show
  End If
End Sub

Private Sub MIngBodega_Click()
  RatonReloj
  IngBodega.Show
End Sub

Private Sub MIngCosteo_Click()
  RatonReloj
  Kard_Ing_DYE.Show
End Sub

Private Sub MIngMarcas_Click()
  RatonReloj
  IngMarcas.Show
End Sub

Private Sub MIngMercDUI_Click()
  RatonReloj
  Kard_Ing_DUI.Show
End Sub

Private Sub MIngProv_Click()
  RatonReloj
  FClientes.Show
End Sub

Private Sub MInvProdTerminados_Click()
  RatonReloj
  Ing_Combo.Show
End Sub

'''Private Sub MIngStockArchivo_Click()
'''  RatonReloj
'''  FStockInventario.Show
'''End Sub

'''Private Sub MMantenimiento_Click()
'''  If ClaveAdministrador Then
'''     RatonReloj
'''     FSeteos.Show
'''  End If
'''End Sub

Private Sub MMayorProd_Click()
  Control_Procesos Normal, "Mayorizar Inventario"
  RatonReloj
  sSQL = "UPDATE Trans_Kardex " _
       & "SET Procesado = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  'Ejecutar_SQL_SP sSQL
  RatonReloj
  Mayorizar_Inventario_SP
  'Mayorizar_Inventario
  'MayorizarInv.Show
  RatonNormal
End Sub

Private Sub MMercaderiaPP_Click()
  RatonReloj
  Costos_Recetas.Show
End Sub

Private Sub MReindexar_Kardex_Click()
  Control_Procesos Normal, "Reindexar Inventario"
  RatonReloj
  sSQL = "UPDATE Trans_Kardex " _
       & "SET Procesado = 0 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  RatonReloj
  Mayorizar_Inventario_SP
  RatonNormal
End Sub

Private Sub MResumenTotales_Click()
  RatonReloj
  TotalKardex.Show
End Sub

Private Sub MSalidaCero_Click()
  RatonReloj
  Empleados = False
  Kard_Ing_Ventas.Show
End Sub

Private Sub MSalidaPVP_Click()
  RatonReloj
  SalidaPVP.Show
End Sub

Private Sub MTransf_Bodegas_Click()
  RatonReloj
  FTransferencia_Bodegas.Show
End Sub

Private Sub ResumExist_Click()
  RatonReloj
  ResumenKardex.Show
End Sub

Private Sub salinvart_Click()
  RatonReloj
  KardexSQLs.Show
End Sub

Private Sub SalirS_Click()
  Control_Procesos "Q", "Salir Modulo de Inventario"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  Recordar_Tarea_Hora
End Sub
