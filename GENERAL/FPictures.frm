VERSION 5.00
Begin VB.Form FPictures 
   AutoRedraw      =   -1  'True
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   49.64
   ScaleMode       =   0  'User
   ScaleWidth      =   22
End
Attribute VB_Name = "FPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
  Unload FPictures
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload FPictures
End Sub

Private Sub Form_Load()
   RatonReloj
   AnchoAltoForm FPictures
   FPictures.ScaleMode = vbCentimeters
   FPictures.Width = 20 * 567
   FPictures.Height = 30 * 567
   RutaDestino = RutaSistema & "\FORMATOS\RETENCIO.GIF"
   FPictures.PaintPicture LoadPicture(RutaDestino), 1, 0.1, 18.5, 20
   RatonNormal
End Sub

Private Sub Image1_Click()
End Sub
