VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPuntoVenta 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Facturacion"
   ClientHeight    =   6390
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7245
   Icon            =   "MDIPVent.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPVent.frx":0B9A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7245
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
      Top             =   6015
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIPVent.frx":26D58
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIPVent.frx":273E2
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "MDIPVent.frx":27CBC
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIPVent.frx":27FD6
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPVent.frx":288B0
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
      Top             =   5895
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu BaseRel 
      Caption         =   "&Archivos"
      Begin VB.Menu MDelSistema 
         Caption         =   "Del Sistema"
         Begin VB.Menu MModAud 
            Caption         =   "&Modulos de Auditoria"
         End
      End
      Begin VB.Menu MClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Productos 
         Caption         =   "Ingresar &Productos (Ventas)"
      End
      Begin VB.Menu MBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MIngFactHotel 
         Caption         =   "Punto de Venta (Facturas)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MPuntoVenta 
         Caption         =   "Punto de Venta (Nota de Venta)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MPuntoVentaTicket 
         Caption         =   "Punto de Venta (Ticket)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MOrdenProduccion 
         Caption         =   "Orden de Produccion"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MBAbono1 
         Caption         =   "-"
      End
      Begin VB.Menu ListarFact 
         Caption         =   "Listar/Anular Punto de Ventas"
         Shortcut        =   ^L
      End
      Begin VB.Menu MBAbono2 
         Caption         =   "-"
      End
      Begin VB.Menu MCierredelPV 
         Caption         =   "Cierre Diario de Caja de Tickes/&Puntos de Venta"
         Shortcut        =   {F5}
      End
      Begin VB.Menu CobDia 
         Caption         =   "Cierre del &Diario de Caja"
         Shortcut        =   {F6}
      End
      Begin VB.Menu SalirSyst 
         Caption         =   "-"
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
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "H"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CobDia_Click()
  RatonReloj
  FCierreCaja.Show
End Sub

Private Sub ListarFact_Click()
   RatonReloj
   ListFact.Show
End Sub

Private Sub MCatProd_Click()
  RatonReloj
  CatalogoCtas.Show
End Sub

Private Sub MCierredelPV_Click()
  RatonReloj
  FConvertirPV.Show
End Sub

Private Sub MClientes_Click()
  RatonReloj
  FClientes.Show
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = MDIFormulario.ScaleHeight - 100
  MDI_X_Max = MDIFormulario.ScaleWidth - 100
End Sub

Private Sub MDIForm_Load()
  Set MDIFormulario = Me
  Primera_Vez = True
  Cyber_Cabinas = True
  TipoFactura = "PV"
  UnidadSistema
  TipoModulo = Factu
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "CYBER NET"
  MenuDeModulos = True
  TiempoSistema = Time
  Timer1.Interval = 1000
  IngresarClave = True
  ListEmp.Show 1
  PonerDirEmpresa
 'Ver_Grafico_FormPict MDIPuntoVenta
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Punto de Venta"
  End
End Sub

Private Sub MIngFactHotel_Click()
  RatonReloj
  Cyber_Cabinas = True
  TipoFactura = "FA"
  FacturaNueva = True
  FPuntoVenta.Show
End Sub

Private Sub MModAud_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Ingreso al Modulo de Auditoria"
     FAuditoria.Show
  End If
End Sub

Private Sub MOrdenProduccion_Click()
  RatonReloj
  Cyber_Cabinas = True
  TipoFactura = "OP"
  FacturaNueva = True
  FPuntoVenta.Show
End Sub

Private Sub MPuntoVenta_Click()
  RatonReloj
  Cyber_Cabinas = True
  TipoFactura = "NV"
  FacturaNueva = True
  FPuntoVenta.Show
End Sub

Private Sub MPuntoVentaTicket_Click()
  RatonReloj
  Cyber_Cabinas = True
  TipoFactura = "PV"
  FacturaNueva = True
  FPuntoVenta.Show
End Sub

Private Sub Productos_Click()
   RatonReloj
   IngProdInv.Show
End Sub

Private Sub SalirSystem_Click()
  Control_Procesos "Q", "Salir Modulo de Punto de Venta"
  End
End Sub

Private Sub Timer1_Timer()
'''  If Len(NumEmpresa) > 1 Then
'''     sSQL = "UPDATE Catalogo_Cyber " _
'''          & "SET Inicio = '00:00:00' " _
'''          & "WHERE Item = '" & NumEmpresa & "' " _
'''          & "AND Periodo = '" & Periodo_Contable & "' " _
'''          & "AND PC_Ocupaga = " & Val(adFalse) & " "
'''     ConectarAdoExecute sSQL
'''  End If
Dim MiTiempoFin As Single
Dim IPC As Integer
Dim AdoCyber As ADODB.Recordset
''  Parpadear = Not Parpadear
''  If Len(NumEmpresa) >= 3 And Cyber_Cabinas Then
''     Set AdoCyber = New ADODB.Recordset
''     AdoCyber.CursorType = adOpenStatic
''     AdoCyber.CursorLocation = adUseClient
''
''     For IPC = 0 To 7
''         TiempoPCs(IPC) = "00:00:00"
''         TotalPCs(IPC) = 0
''     Next IPC
''     MiTiempoFin = CDbl(CDate(Time))
''     sSQL = "SELECT * " _
''          & "FROM Catalogo_Cyber " _
''          & "WHERE Item = '" & NumEmpresa & "' " _
''          & "AND Periodo = '" & Periodo_Contable & "' " _
''          & "ORDER BY Codigo "
''     CompilarSQL sSQL
''     AdoCyber.Open sSQL, AdoStrCnn, , , adCmdText
''     With AdoCyber
''      If .RecordCount Then
''          Do While Not .EOF
''             Codigo = .Fields("Codigo")
''             Codigo = Mid$(Codigo, Len(Codigo) - 1, 2)
''             IPC = Val(Codigo) - 1
''             MiTiempo = CDbl(CDate(.Fields("Inicio")))
''            'MsgBox Codigo & vbCrLf & IPC
''             If .Fields("PC_Ocupaga") Then
''                 Calcular_Total_PC IPC, MiTiempo, MiTiempoFin
''                 If Parpadear Then
''                    FCyberNet.CButtonPCs(IPC).ForeColor = RGB(&HDC, &H14, &H3C)
''                 Else
''                    FCyberNet.CButtonPCs(IPC).ForeColor = RGB(&H94, &H0, &HD3)
''                 End If
''                 FCyberNet.CButtonPCs(IPC).Caption = " Equipo " & IPC + 1 _
''                          & " [" & Format(TiempoPCs(IPC), FormatoTimes) & "] " _
''                          & "USD " & TotalPCs(IPC)
''             End If
''            .MoveNext
''          Loop
''      End If
''     End With
''     sSQL = "UPDATE Catalogo_Cyber " _
''          & "SET Fin = '" & CDate(Time) & "' " _
''          & "WHERE Item = '" & NumEmpresa & "' " _
''          & "AND Periodo = '" & Periodo_Contable & "' "
''     ConectarAdoExecute sSQL
''  End If
  Ver_Grafico_FormPict
End Sub
