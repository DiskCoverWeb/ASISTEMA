VERSION 5.00
Begin VB.Form FConexion 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtConexion 
      BackColor       =   &H00FFC0C0&
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
      Height          =   1380
      Left            =   735
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FConexion.frx":0000
      Top             =   210
      Width           =   6840
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   105
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   0
      Left            =   105
      Picture         =   "FConexion.frx":0017
      Top             =   210
      Width           =   510
   End
End
Attribute VB_Name = "FConexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nFrames As Long
Dim AnchoMaxForm As Single
    RatonReloj
    CentrarForm FConexion
    'TxtConexion = "CONECTANDO AL SRI..."
    'TxtConexion.Refresh
    Redondear_Formulario FConexion, 40
        
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
TxtConexion.ForeColor = Azul   ' &HC00000
If Err Then Exit Sub
End Sub

