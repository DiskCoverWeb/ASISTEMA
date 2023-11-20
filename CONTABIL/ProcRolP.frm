VERSION 5.00
Begin VB.Form FProcesarRolPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PROCESO DE ROL DE PAGOS (NOMINA)"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FProcesarRolPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoNomina
  ConectarAdodc AdoAux
  RatonNormal
End Sub
