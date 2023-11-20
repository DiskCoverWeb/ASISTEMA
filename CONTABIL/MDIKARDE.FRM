VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIKardex 
   BackColor       =   &H8000000C&
   Caption         =   "MDI: KARDEK"
   ClientHeight    =   3510
   ClientLeft      =   420
   ClientTop       =   270
   ClientWidth     =   6570
   Icon            =   "MDIKarde.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3180
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Picture         =   "MDIKarde.frx":0442
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu BaseReferencial 
      Caption         =   "Bases Referenciales"
      Begin VB.Menu TablaBase 
         Caption         =   "Del Sistema"
      End
      Begin VB.Menu xx 
         Caption         =   "-"
      End
      Begin VB.Menu xxx 
         Caption         =   "De Procesos"
         Begin VB.Menu Productos 
            Caption         =   "Ingreso de Productos"
         End
         Begin VB.Menu IngProd 
            Caption         =   "Ingreso de Unidades de Medidas"
         End
         Begin VB.Menu Invent 
            Caption         =   "Inventario"
            Begin VB.Menu IngMatPrim 
               Caption         =   "Ingresos"
            End
            Begin VB.Menu EgreMatPrim 
               Caption         =   "Egresos"
            End
         End
      End
   End
   Begin VB.Menu REportes 
      Caption         =   "Reportes"
      Begin VB.Menu KardexProd 
         Caption         =   "KARDEX por producto"
      End
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
      Begin VB.Menu SalirSyst 
         Caption         =   "Salir del Sistema"
         Shortcut        =   ^Q
      End
      Begin VB.Menu CambiarEmp 
         Caption         =   "Cambiar de Empresa"
      End
   End
End
Attribute VB_Name = "MDIKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CambiarEmp_Click()
  UnidadSistema
  ListEmp.Show
End Sub

Private Sub IngMatPrim_Click()
  Kard_Ing.Show
End Sub

Private Sub IngProd_Click()
  Ing_Unid.Show
End Sub

Private Sub KardexProd_Click()
  KardexSQLs.Show
End Sub

Private Sub MDIForm_Activate()
  UnidadSistema
  IngClaves.Show
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  MDIKardex.Caption = "KARDEX: " & EmpresaActual
  StatusBar.Panels(1).Text = "Usuario: " & NombreUsuario
End Sub

Private Sub Productos_Click()
   IngProductos.Show
End Sub

Private Sub SalirSyst_Click()
  End
End Sub

Private Sub TablaBase_Click()
  Seteos.Show
End Sub
