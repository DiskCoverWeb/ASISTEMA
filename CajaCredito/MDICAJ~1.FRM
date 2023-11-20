VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDICajaCredito 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7020
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu AperturaCta 
         Caption         =   "Apertura de Cuenta"
      End
      Begin VB.Menu Cajas 
         Caption         =   "Debitos / Creditos"
      End
      Begin VB.Menu FlujoCaja 
         Caption         =   "Flujo de Caja"
      End
      Begin VB.Menu SaldoDiarios 
         Caption         =   "Saldos Diarios"
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu Impresoras 
         Caption         =   "Impresora"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Resportes 
      Caption         =   "Reportes"
   End
   Begin VB.Menu Programador 
      Caption         =   "Programador"
   End
End
Attribute VB_Name = "MDICajaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AperturaCta_Click()
  RatonReloj
  Apertura.Show
End Sub

Private Sub Cajas_Click()
   RatonReloj
   FCaja.Show
End Sub

Private Sub FlujoCaja_Click()
  RatonReloj
  FlujoDeCaja.Show
End Sub

Private Sub MDIForm_Load()
  UnidadSistema
  IngresarClave = True
 'MODULOS
  'MenuDeModulos = True
  ListEmp.Show
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PonerDirEmpresa MDICajaCredito, "CAJA CREDITO"
  'If Supervisor = False Then
  '   DelSyst.Enabled = CNivel_1
  '   Cuentas.Enabled = CNivel_2
  '   IngSubCtasBloq.Enabled = CNivel_2
  '   CompIngresos.Enabled = CNivel_3
  '   CompEgresos.Enabled = CNivel_3
  '   Regulaciones.Enabled = CNivel_3
  '   MConciliar.Enabled = CNivel_3
  '   Mayorizar.Enabled = CNivel_4
  '   ProcesarBal1.Enabled = CNivel_4
  '   BalCon1.Enabled = CNivel_5
  '   BalanceSitua.Enabled = CNivel_5
  '   EstadoResult.Enabled = CNivel_5
  '   AnularComp.Enabled = CNivel_6
  'End If
End Sub

Private Sub Programador_Click()
   RatonReloj
   PagPrint.Show
End Sub

Private Sub SaldoDiarios_Click()
  RatonReloj
  SaldoDiarioCtas.Show
End Sub

Private Sub Salir_Click()
   End
End Sub
