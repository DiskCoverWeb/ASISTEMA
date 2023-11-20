VERSION 5.00
Begin VB.Form FBusqueda 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4095
      Picture         =   "FBusq.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.TextBox TxtBusqueda 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   3900
   End
   Begin VB.Label LblBusqueda 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Datos de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3900
   End
End
Attribute VB_Name = "FBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
  TextoBusqueda = ""
  Unload FBusqueda
End Sub

Private Sub Form_Activate()
  RatonNormal
  FBusqueda.Visible = True
End Sub

Private Sub Form_Load()
  CentrarForm FBusqueda
  LblBusqueda.Caption = " Campo: " & CampoBusqueda
End Sub

Private Sub TxtBusqueda_GotFocus()
  MarcarTexto TxtBusqueda
End Sub

Private Sub TxtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload FBusqueda
  If KeyCode = vbKeyReturn Then
     TextoValido TxtBusqueda
     If TipoDatoBusq <> 0 Then
        Select Case TipoDatoBusq
          Case TadDate, TadDate1
               TextoBusqueda = " like '" & TxtBusqueda.Text & "' "
          Case TadByte, TadInteger, TadLong, TadSingle, TadDouble, TadCurrency, TadBoolean
               TextoValido TxtBusqueda, True
               TextoBusqueda = " like " & TxtBusqueda.Text & " "
          Case TadText, TadMemo
               TextoBusqueda = " like '" & TxtBusqueda.Text & "*' "
          Case Else
               TextoBusqueda = " like '" & TxtBusqueda.Text & "*' "
        End Select
     Else
        MsgBox "No existe Datos que buscar"
     End If
  Unload FBusqueda
  End If
End Sub
