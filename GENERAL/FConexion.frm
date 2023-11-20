VERSION 5.00
Begin VB.Form FConexion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   60
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtConexion 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   735
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FConexion.frx":0000
      Top             =   105
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   0
      Left            =   105
      Picture         =   "FConexion.frx":0017
      Top             =   105
      Width           =   510
   End
End
Attribute VB_Name = "FConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim Tiempo_Espera As Integer
    For Tiempo_Espera = 0 To 600
        TxtConexion = "CONECTANDO AL SRI..."
    Next Tiempo_Espera
End Sub

Private Sub Form_Load()
Dim nFrames As Long
Dim AnchoMaxForm As Single
    RatonReloj
    nFrames = Load_Gif(RutaSistema & "\FORMATOS\conexion.gif", Image1)
    If nFrames > 0 Then
       FrameCount = 0
       Timer1.Interval = 2000
       Timer1.Enabled = True
    End If
    AnchoMaxForm = FConexion.width
    If AnchoMaxForm < FConexion.TextWidth(Progreso_Barra.Mensaje_Box) Then
       AnchoMaxForm = FConexion.TextWidth(Progreso_Barra.Mensaje_Box)
    End If
    AnchoMaxForm = AnchoMaxForm + 1500
    FConexion.width = AnchoMaxForm
    FConexion.TxtConexion.width = AnchoMaxForm - 1000
    CentrarForm FConexion
    Redondear_Formulario FConexion, 30
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim I As Long
If FrameCount < TotalFrames Then
   Image1(FrameCount).Visible = False
   FrameCount = FrameCount + 1
Else
   FrameCount = 0
   For I = 1 To Image1.Count - 1
   Image1(I).Visible = False
   Next I
End If

Image1(FrameCount).Visible = True
Timer1.Interval = CLng(Image1(FrameCount).Tag)
TxtConexion.ForeColor = Verde   ' &HC00000
If Err Then Exit Sub
End Sub

