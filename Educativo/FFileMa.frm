VERSION 5.00
Begin VB.Form FFileMaterias 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   19.42
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.882
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "FFileMaterias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Activate()
  Si_No = ExisteFormulario("A")
  RatonNormal
End Sub

Private Sub Form_Load()
End Sub
