VERSION 5.00
Begin VB.Form FBusquedaList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSCAR DATOS"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   5040
      Picture         =   "FBusqL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4095
      Picture         =   "FBusqL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   315
      Width           =   855
   End
   Begin VB.TextBox TxtBusqueda 
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
      TabIndex        =   1
      Top             =   630
      Width           =   3900
   End
   Begin VB.Label LblBusqueda 
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
      Top             =   315
      Width           =   3900
   End
End
Attribute VB_Name = "FBusquedaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If TipoDatoBusq <> 0 Then
     Select Case TipoDatoBusq
       Case TadDate, TadDate1
            TextoBusqueda = " like '" & TxtBusqueda.Text & "' "
       Case TadByte, TadInteger, TadLong, TadSingle, TadDouble, TadBoolean
            TextoBusqueda = " like " & TxtBusqueda.Text & " "
       Case TadText, TadMemo
            TextoBusqueda = " like '" & TxtBusqueda.Text & "' "
       Case Else
            TextoBusqueda = " like '" & TxtBusqueda.Text & "' "
     End Select
  Else
     MsgBox "No existe Datos que buscar"
  End If
  Unload FBusqueda
End Sub

Private Sub Command2_Click()
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
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload FBusqueda
End Sub
