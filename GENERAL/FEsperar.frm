VERSION 5.00
Begin VB.Form FEsperar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   885
   ClientLeft      =   15
   ClientTop       =   30
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5670
      Top             =   210
   End
   Begin VB.Label LblMensaje 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CONECTANDO A LA BASE DE DATOS..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   945
      TabIndex        =   0
      Top             =   105
      Width           =   5160
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   105
      Picture         =   "FEsperar.frx":0000
      Stretch         =   -1  'True
      Top             =   105
      Width           =   600
   End
End
Attribute VB_Name = "FEsperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nFrames As Long
Dim Tiempo_Espera As Integer
Dim RutaGifs As String
    RatonReloj
    CentrarForm FEsperar
    Redondear_Formulario FEsperar, 35
    
   'Nombre_Imagen_Esperar = "conexion"
    Image1.Item(0).Height = 600
    Image1.Item(0).width = 600
    RutaGifs = RutaSistema & "\FORMATOS\" & NombreImagenEsperar & ".gif"
    If Dir$(RutaGifs) = "" Then RutaGifs = RutaSistema & "\FORMATOS\procesando.gif"
    nFrames = Load_Gif(RutaGifs, Image1)
    Timer1.Interval = 120
    Timer1.Enabled = True
    If nFrames > 0 Then
       For Tiempo_Espera = 0 To nFrames
           Image1(Tiempo_Espera).Visible = True
           Image1(Tiempo_Espera).Refresh
       Next Tiempo_Espera
       Sleep 5
       For Tiempo_Espera = 0 To nFrames - 1
           Image1(Tiempo_Espera).Visible = False
           Image1(Tiempo_Espera).Refresh
           Sleep 5
           Image1(Tiempo_Espera + 1).Visible = True
           Image1(Tiempo_Espera + 1).Refresh
           Sleep 5
       Next Tiempo_Espera
       FrameCount = 0
    End If
    LblMensaje.Left = Image1(0).Left + Image1(0).width + 150
    LblMensaje.Refresh
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NombreImagenEsperar = "procesando"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim I As Long

'MsgBox TotalFrames
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
Timer1.Interval = 120 ' CLng(Image1(FrameCount).Tag)
FEsperar.width = LblMensaje.Left + LblMensaje.width + 100
LblMensaje.Refresh
'Label1.ForeColor = Verde   ' &HC00000
If Err Then Exit Sub
End Sub

