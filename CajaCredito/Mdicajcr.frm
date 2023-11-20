VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDICajaCredito 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDI"
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7650
   Icon            =   "Mdicajcr.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Mdicajcr.frx":1CCA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7650
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7650
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   105
      Top             =   105
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "Mdicajcr.frx":27E88
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "Mdicajcr.frx":28512
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Picture         =   "Mdicajcr.frx":28DEC
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "Mdicajcr.frx":29106
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "Mdicajcr.frx":299E0
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
      Top             =   6120
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Menu Archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MOpSys 
         Caption         =   "Del &Sistema"
         Begin VB.Menu MCorrCred 
            Caption         =   "&Corrección de Créditos"
         End
         Begin VB.Menu MCambPeriodo 
            Caption         =   "Cambio de &Periodo"
         End
         Begin VB.Menu MMigracio 
            Caption         =   "Migracion de Datos"
         End
         Begin VB.Menu MCatCreditos 
            Caption         =   "Ingresar Catalogo de Créditos"
         End
      End
      Begin VB.Menu MDeOper 
         Caption         =   "De Operación"
         Begin VB.Menu AperturaCta 
            Caption         =   "&Apertura de Cuenta"
         End
         Begin VB.Menu MopeCaja 
            Caption         =   "&Operaciones de Caja"
            Begin VB.Menu Cajas 
               Caption         =   "&Depósitos / Retiros"
               Shortcut        =   ^M
            End
            Begin VB.Menu MEfecCheq 
               Caption         =   "&Efectivizar Cheques"
            End
            Begin VB.Menu SaldoDiarios 
               Caption         =   "&Intereses Diarios"
            End
            Begin VB.Menu MAsientCont 
               Caption         =   "&Cierre de Flujo de Caja"
               Shortcut        =   ^F
            End
            Begin VB.Menu MAcredInt 
               Caption         =   "Acreditar I&ntereses"
            End
            Begin VB.Menu MImportarTablaDep 
               Caption         =   "Importar Tabla de Depositos"
            End
         End
         Begin VB.Menu MOpeCredito 
            Caption         =   "O&peraciones de Credito"
            Begin VB.Menu MPrestamos 
               Caption         =   "&Liquidacion de Prestamos"
            End
            Begin VB.Menu MAprobCred 
               Caption         =   "&Validacion del Credito"
            End
            Begin VB.Menu MAbonosVencidos 
               Caption         =   "Abonos/Cancelación de &Vencidos"
            End
            Begin VB.Menu MPrecancPrest 
               Caption         =   "&Precancelacion de Préstamos"
            End
            Begin VB.Menu TrabsCred 
               Caption         =   "&Notas de: Debitos/Creditos/Certificados"
            End
            Begin VB.Menu MMovTarjeta 
               Caption         =   "&Movimientos de Tarjetas"
            End
            Begin VB.Menu FTransGrupLib 
               Caption         =   "&Costo por Mantenimiento/Fondo Mortuorio"
            End
            Begin VB.Menu MEncaje 
               Caption         =   "&Encaje o Bloqueo de Retiros"
            End
            Begin VB.Menu MAbonoPrestAntiguo 
               Caption         =   "Abonos Prestamos Antiguos"
            End
         End
         Begin VB.Menu MAsigCobrosAuto 
            Caption         =   "Asignación de Cobros Automáticos"
         End
         Begin VB.Menu MPolizas 
            Caption         =   "Inversiones a Corto Plazo"
         End
      End
      Begin VB.Menu FlujoCaja 
         Caption         =   "Flujo de &Caja"
      End
      Begin VB.Menu Fludelib 
         Caption         =   "Flujo de &Libretas"
      End
      Begin VB.Menu MFlujoPrestamos 
         Caption         =   "Flujo de &Prestamos"
      End
      Begin VB.Menu MFlujoSusp 
         Caption         =   "Flujo de &Suspenso/Cheques"
      End
      Begin VB.Menu MFlujoDepEdu 
         Caption         =   "Flujo Depositos Educativos"
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu MCambEmp 
         Caption         =   "Cambiar de &Empresa"
         Shortcut        =   ^E
      End
      Begin VB.Menu Salir 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Resportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MListCtas 
         Caption         =   "Listar Cuentas"
      End
      Begin VB.Menu SaldoLib 
         Caption         =   "Saldo de Libretas"
      End
      Begin VB.Menu MSaldoPrest 
         Caption         =   "Saldo de Prestamos"
      End
      Begin VB.Menu MListMovCtas 
         Caption         =   "Historial Cuentas/Depositos Cheques"
         Shortcut        =   ^H
      End
      Begin VB.Menu MCreditosOtorgados 
         Caption         =   "Creditos Otorgados"
      End
      Begin VB.Menu MVencidos 
         Caption         =   "Creditos Vencidos"
         Shortcut        =   ^P
      End
      Begin VB.Menu Mb2 
         Caption         =   "-"
      End
      Begin VB.Menu MConsultaAbonos 
         Caption         =   "Consulta de Abonos"
      End
      Begin VB.Menu MListGarantes 
         Caption         =   "Listado de Garantes"
      End
      Begin VB.Menu MRepUAF 
         Caption         =   "-"
      End
      Begin VB.Menu MReportesUAF 
         Caption         =   "Reportes a la UAF"
      End
   End
   Begin VB.Menu MProcVarios 
      Caption         =   "Procesos Varios"
      Begin VB.Menu Programador 
         Caption         =   "Programador"
      End
      Begin VB.Menu MReProcSaldo_Lib 
         Caption         =   "Reprocesar Saldos Libretas"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "H"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDICajaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AperturaCta_Click()
  RatonReloj
  FClientes.Show
End Sub

Private Sub Cajas_Click()
   RatonReloj
   OpcCaja = True
   FCaja.Show
End Sub

Private Sub Fludelib_Click()
  RatonReloj
  FlujoDeLibretas.Show
End Sub

Private Sub FlujoCaja_Click()
  RatonReloj
  FlujoDeCaja.Show
End Sub

Private Sub FTransGrupLib_Click()
  RatonReloj
  ProcesarND_NC.Show
End Sub

Private Sub MAbonoPrestAntiguo_Click()
  RatonReloj
  AbonoPrestamoManual.Show
End Sub

Private Sub MAbonosVencidos_Click()
  RatonReloj
  AbonoPrestamo.Show
End Sub

Private Sub MAcredInt_Click()
 RatonReloj
 FIntLibretas.Show
End Sub

Private Sub MAprobCred_Click()
  RatonReloj
  Aprobacion.Show
End Sub

Private Sub MAsientCont_Click()
   TipoNumAsiento = 1
   FAsientos.Show
End Sub

Private Sub MAsigCobrosAuto_Click()
  RatonReloj
  ListarGrupos.Show
End Sub

Private Sub MCambEmp_Click()
  RatonReloj
  UnidadSistema
  IngresarClave = False
  ListEmp.Show
End Sub

Private Sub MCambPeriodo_Click()
  If ClaveContador Then
     RatonReloj
     Periodos.Show
  End If
End Sub

Private Sub MCatCreditos_Click()
  RatonReloj
  IngPlanCreditos.Show
End Sub

Private Sub MConsultaAbonos_Click()
  RatonReloj
  ListarAbonosPrest.Show
End Sub

Private Sub MCorrCred_Click()
  If ClaveAuxiliar Then
     RatonReloj
     CorreccionCredito.Show
  End If
End Sub

Private Sub MCreditosOtorgados_Click()
  RatonReloj
  RPrestamo.Show
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
  TipoModulo = conta
  IngresarClave = True
 'MODULOS
  NumModulo = "0"
  Modulo = "CAJACREDITO"
  MenuDeModulos = True
 'TiempoTarea = Time
  TiempoSistema = Time
  Timer1.Enabled = True
  Timer1.Interval = 1000
  ListEmp.Show 1
  PonerDirEmpresa
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Caja Credito"
  End
End Sub

Private Sub MEfecCheq_Click()
  RatonReloj
  ChequeEfectivo.Show
End Sub

Private Sub MEncaje_Click()
   RatonReloj
   Encaje.Show
End Sub

Private Sub MFlujoDepEdu_Click()
  RatonReloj
  FAbonos.Show
End Sub

Private Sub MFlujoPrestamos_Click()
  RatonReloj
  FlujoDePrestamos.Show
End Sub

Private Sub MFlujoSusp_Click()
  RatonReloj
  FlujoSuspenso.Show
End Sub

Private Sub MImportarTablaDep_Click()
  RatonReloj
  FImporta.Show
End Sub

Private Sub MListCtas_Click()
  RatonReloj
  FListarCtas.Show
End Sub

Private Sub MListGarantes_Click()
  RatonReloj
  FGarantes.Show
End Sub

Private Sub MListMovCtas_Click()
  RatonReloj
  Historial.Show
End Sub

'''Private Sub MMantenimiento_Click()
'''  If ClaveSupervisor Then
'''     RatonReloj
'''     FSeteos.Show
'''  End If
'''End Sub

Private Sub MMigracio_Click()
  If ClaveAdministrador Then
     RatonReloj
     FMigracion.Show
  End If
End Sub

Private Sub MMovTarjeta_Click()
  RatonReloj
  FTarjeta.Show
End Sub

Private Sub MPolizas_Click()
  RatonReloj
  FPolizas.Show
End Sub

Private Sub MPrecancPrest_Click()
  RatonReloj
  Precancelacion.Show
End Sub

Private Sub MPrestamos_Click()
   RatonReloj
   FPrestamo.Show
End Sub

Private Sub MReportesUAF_Click()
  RatonReloj
  FReportesUAF.Show
End Sub

Private Sub MReProcSaldo_Lib_Click()
Dim Num_Comp As Long
Dim CantCtas As Integer
Dim CantSubCtas As Integer

Dim Primero As Boolean
Dim Si_No_SC As Boolean
Dim Cod_Cta As String
Dim ValorStr As String
Dim SubModulos() As String
Dim TempNumEmpresa As String

Dim AdoCtasDB As ADODB.Recordset
Dim AdoTransDB As ADODB.Recordset
 
   RatonReloj
   Control_Procesos Normal, "Mayorizar Cuentas"
  
  CantCtas = 100
  TextoImprimio = ""
 
 'INICIO DE LA MAYORIZACION
  Progreso_Barra.Mensaje_Box = "Iniciando la mayorizacion Ctas. Espere un momento..."
  Progreso_Iniciar
  
  sSQL = "UPDATE Clientes_Datos_Extras " _
       & "SET Procesado = 1 " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' "
  ConectarAdoExecute sSQL
  
    If SQL_Server Then
       sSQL = "UPDATE Clientes_Datos_Extras " _
            & "SET Procesado = 0 " _
            & "FROM Clientes_Datos_Extras As CL,Trans_Libretas As TL "
    Else
       sSQL = "UPDATE Clientes_Datos_Extras As CL,Trans_Libretas As TL " _
            & "SET CL.Procesado = 0 "
    End If
    sSQL = sSQL _
         & "WHERE TL.Item = '" & NumEmpresa & "' " _
         & "AND CL.Tipo_Dato = 'LIBRETAS' " _
         & "AND TL.Procesado = " & Val(adFalse) & " " _
         & "AND CL.Cuenta_No = TL.Cuenta_No " _
         & "AND CL.Item = TL.Item "
    ConectarAdoExecute sSQL
  
  sSQL = "SELECT Cuenta_No " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Tipo_Dato = 'LIBRETAS' " _
       & "AND Procesado = " & Val(adFalse) & " " _
       & "ORDER BY Cuenta_No "
  SelectAdoDB AdoCtasDB, sSQL
  CantCtas = CantCtas + AdoCtasDB.RecordCount
  
  Progreso_Barra.Valor_Maximo = CantCtas
  Progreso_Esperar
  If AdoCtasDB.RecordCount > 0 Then
     Do While Not AdoCtasDB.EOF
       'Mayoriazar cuentas contables
        Cod_Cta = AdoCtasDB.Fields("Cuenta_No")
        Progreso_Barra.Mensaje_Box = "Mayorizando Cuenta No. " & Cod_Cta & "..."
        Progreso_Esperar
        sSQL = "SELECT * " _
             & "FROM Trans_Libretas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Cuenta_No = '" & Cod_Cta & "' " _
             & "ORDER BY Fecha,Hora,ID "
        SelectAdoDB AdoTransDB, sSQL
        RatonReloj
        SumaDebe = 0: SumaHaber = 0
        With AdoTransDB
         If .RecordCount > 0 Then
             CantSubCtas = .RecordCount
             FechaTexto = .Fields("Fecha")
             Debe = .Fields("Debitos")
             Haber = .Fields("Creditos")
             Saldo = 0
             Mifecha = .Fields("Fecha")
             NoMes = Month(.Fields("Fecha"))
             Contador = 0
             Do While Not .EOF
                Mifecha = .Fields("Fecha")
                Debe = Redondear(.Fields("Debitos"), 2)
                Haber = Redondear(.Fields("Creditos"), 2)
                Saldo = Saldo + Haber - Debe
                Progreso_Barra.Mensaje_Box = Format(Contador / CantSubCtas, "00%") & " Mayorizando Cuenta No. " & Cod_Cta & " Fecha: " & Mifecha
               .Fields("Saldo_Lib") = Saldo
               .Fields("Procesado") = True
               .Update
               'MsgBox ProcBar.Value
                Progreso_Esperar True
                Contador = Contador + 1
               .MoveNext
             Loop
         End If
        End With
        AdoTransDB.Close
       'Siguiente Cta
       'msgbox Cod_Cta
        AdoCtasDB.MoveNext
     Loop
  Else
     Progreso_Barra.Mensaje_Box = "No existe Cuentas a Mayorizar"
     Progreso_Esperar
  End If
  AdoCtasDB.Close
  RatonNormal
  Progreso_Barra.Mensaje_Box = "PROCESO TERMINADO"
  Progreso_Final
  
  If TextoImprimio <> "" Then FInfoError.Show
End Sub

Private Sub MSaldoPrest_Click()
  RatonReloj
  SaldoPrestamo.Show
End Sub

Private Sub MVencidos_Click()
  RatonReloj
  FVencidos.Show
End Sub

Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
   'FSocket.Show
End Sub

Private Sub SaldoDiarios_Click()
  RatonReloj
  FIntereses.Show
End Sub

Private Sub SaldoLib_Click()
  RatonReloj
  FListarSaldoCtas.Show
End Sub

Private Sub Salir_Click()
  Control_Procesos "Q", "Salir Modulo de Caja Credito"
  End
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
  If Supervisor = False And Len(CodigoUsuario) > 1 Then
   ' Seteamos los menus
     Cajas.Enabled = False
     MEfecCheq.Enabled = False
     SaldoDiarios.Enabled = False
     MAsientCont.Enabled = False
     MAcredInt.Enabled = False
     AperturaCta.Enabled = False
     MPrestamos.Enabled = False
     MAprobCred.Enabled = False
     MAbonosVencidos.Enabled = False
     MPrecancPrest.Enabled = False
     TrabsCred.Enabled = False
     FTransGrupLib.Enabled = False
     MEncaje.Enabled = False
     FlujoCaja.Enabled = False
     Fludelib.Enabled = False
     MListMovCtas.Enabled = False
     MCreditosOtorgados.Enabled = False
     MVencidos.Enabled = False
     MMovTarjeta.Enabled = False
     MFlujoPrestamos.Enabled = False
     MFlujoSusp.Enabled = False
     MListCtas.Enabled = False
     SaldoLib.Enabled = False
   ' Seteamos los menus
     Cajas.Enabled = CNivel(1) Or CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     MEfecCheq.Enabled = CNivel(3)
     SaldoDiarios.Enabled = CNivel(3)
     MAsientCont.Enabled = CNivel(3) Or CNivel(5)
     MAcredInt.Enabled = CNivel(3)
     AperturaCta.Enabled = CNivel(1) Or CNivel(3) Or CNivel(4) Or CNivel(5)
     MPrestamos.Enabled = CNivel(3) Or CNivel(4)
     MAprobCred.Enabled = CNivel(3)
     MAbonosVencidos.Enabled = CNivel(3) Or CNivel(4)
     MPrecancPrest.Enabled = CNivel(3) Or CNivel(4)
     TrabsCred.Enabled = CNivel(2) Or CNivel(4) Or CNivel(5)
     FTransGrupLib.Enabled = CNivel(3)
     MEncaje.Enabled = CNivel(3)
     FlujoCaja.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     Fludelib.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     MListMovCtas.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     MCreditosOtorgados.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(6)
     MVencidos.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(6)
     MMovTarjeta.Enabled = CNivel(1) Or CNivel(3)
     MFlujoPrestamos.Enabled = CNivel(2) Or CNivel(3) Or CNivel(4) Or CNivel(6)
     MFlujoSusp.Enabled = CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     MListCtas.Enabled = CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
     SaldoLib.Enabled = CNivel(3) Or CNivel(4) Or CNivel(5) Or CNivel(6)
  End If
  Recordar_Tarea_Hora
End Sub

Private Sub TrabsCred_Click()
   RatonReloj
   OpcCaja = False
   FCaja.Show
End Sub
