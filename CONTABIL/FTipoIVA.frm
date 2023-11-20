VERSION 5.00
Begin VB.Form FTipoIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo Retencion I.V.A."
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstTipoIVA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2850
   End
End
Attribute VB_Name = "FTipoIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 CentrarForm FTipoIVA
 LstTipoIVA.Clear
 LstTipoIVA.AddItem "1.- Compras Locales"
 LstTipoIVA.AddItem "2.- Ventas Locales"
 LstTipoIVA.AddItem "3.- Importaciones"
 LstTipoIVA.AddItem "4.- Exportaciones"
 LstTipoIVA.Text = LstTipoIVA.List(0)
End Sub

Private Sub LstTipoIVA_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     Topc = Mid(LstTipoIVA.Text, 1, 1)
     Unload Me
     FRetIVA.Show 1
  End If
  If KeyCode = vbKeyEscape Then Unload Me
End Sub
