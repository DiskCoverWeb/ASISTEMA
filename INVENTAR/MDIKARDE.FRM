VERSION 5.00
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
            Begin VB.Menu MatPrim 
               Caption         =   "Materia Prima"
               Begin VB.Menu IngMatPrim 
                  Caption         =   "Ingresos"
               End
               Begin VB.Menu EgreMatPrim 
                  Caption         =   "Egresos"
               End
            End
            Begin VB.Menu ProdProc 
               Caption         =   "Productos en Procesos"
               Begin VB.Menu IngProdProc 
                  Caption         =   "Ingresos"
               End
               Begin VB.Menu EgreProdProc 
                  Caption         =   "Egresos"
               End
            End
            Begin VB.Menu ProdTerm 
               Caption         =   "Productos Terminados"
               Begin VB.Menu IngProdTerm 
                  Caption         =   "Ingresos"
               End
               Begin VB.Menu EgreProdTerm 
                  Caption         =   "Egreso"
               End
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

Private Sub IngKad_Click()
   Kard_Ing.Show
End Sub

Private Sub IngProd_Click()
  Ing_Unid.Show
End Sub

Private Sub Kardek_Click()
  Kard_Ing.Show
End Sub

Private Sub KardexProd_Click()
  KardexSQLs.Show
End Sub

Private Sub MDIForm_Activate()
  UnidadSistema
  ListEmp.Show
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   MDIKardex.Caption = "KARDEX: EMPRESA ACTUAL [ " & Empresa & " ]."
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
