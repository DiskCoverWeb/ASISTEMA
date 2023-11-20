VERSION 5.00
Begin VB.Form FGraficoMatriz 
   Caption         =   "MATRIZ DE RESULTADOS"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13.52
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   14.737
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "&G"
      Height          =   435
      Left            =   2940
      TabIndex        =   3
      Top             =   105
      Width           =   645
   End
   Begin VB.PictureBox PictLetrasV 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4320
      Left            =   2940
      ScaleHeight     =   7.514
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   9.181
      TabIndex        =   2
      Top             =   630
      Width           =   5265
   End
   Begin VB.PictureBox PictLetrasH 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   105
      ScaleHeight     =   2.514
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   4.736
      TabIndex        =   1
      Top             =   105
      Width           =   2745
   End
   Begin VB.PictureBox PictMatriz 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   2430
      Left            =   105
      ScaleHeight     =   4.18
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   16.96
      TabIndex        =   0
      Top             =   5145
      Width           =   9675
   End
End
Attribute VB_Name = "FGraficoMatriz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim PointColor
Dim PictTexto As String
' Empezamos la escritura de las letras
PictLetrasH.AutoRedraw = True
PictLetrasH.Picture = LoadPicture()
PictLetrasH.FontBold = True
PictLetrasH.Font = TipoComicSans
PictLetrasH.FontSize = 10
PictTexto = "Quimestres" & vbCrLf _
          & "Promedio Global" & vbCrLf _
          & "Examen Supletorio" & vbCrLf _
          & "Promedio Total" & vbCrLf
AltoLetra = PictLetrasH.TextHeight(PictTexto)
AnchoMax = PictLetrasH.TextWidth(PictTexto)

PictLetrasH.Visible = False
PictLetrasV.Visible = False
PictLetrasH.width = Redondear(AnchoMax) + 1.5: PictLetrasH.Height = Redondear(AltoLetra) + 0.5

PictLetrasV.Height = PictLetrasH.width: PictLetrasV.width = PictLetrasH.Height

AltoLetra = PictLetrasH.TextHeight(Mid$(PictTexto, 1, 1))
PictLetrasH.Line (AnchoMax + 0.2, 0)-(AnchoMax + 0.2, PictLetrasH.Height), 0
Contador = 0
PCol = 0.1
Msg = ""
PFil = 0.01
For I = 1 To Len(PictTexto)
    If Mid$(PictTexto, I, 1) = vbCr Then
       PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil
       PictLetrasH.Print Msg
       PictLetrasH.Line (0, PFil + 0.7)-(AnchoMax + 0.2, PFil + 0.7), 0
       Msg = ""
       PFil = PFil + 0.3 + AltoLetra
       I = I + 2
    End If
    Msg = Msg & Mid$(PictTexto, I, 1)
Next I
IR = 0
RatonReloj
PictLetrasH.CurrentX = PCol
PictLetrasH.CurrentY = 0.1 + (AltoLetra * Contador)
Do While IR <= PictLetrasH.ScaleWidth
   JR = 0
   Do While JR <= PictLetrasH.ScaleHeight
     PointColor = PictLetrasH.Point(IR, JR)
     PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR - 0.1), PointColor
     JR = JR + 0.005
   Loop
   IR = IR + 0.005
Loop
'PictLetrasH.Image = PictLetrasH.Picture
PictLetrasV.FontBold = True
PictLetrasV.Font = TipoComicSans
PictLetrasV.FontSize = 10
PictLetrasV.CurrentX = 0.1
PictLetrasV.CurrentY = 0.01
PictLetrasV.Print "Lenguaje y "
PictLetrasV.CurrentX = 0.1
PictLetrasV.CurrentY = 0.4
PictLetrasV.Print "Comunicacion"

SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\MATERIAS\PARALELO.BMP"
PictLetrasH.Visible = True
PictLetrasV.Visible = True
RatonNormal
End Sub

Private Sub Form_Activate()

    PictMatriz.Left = 0
    PictMatriz.width = Me.ScaleWidth
    PictMatriz.Height = Me.ScaleHeight - 0.7
     PosLinea = 1
     PosColumna = 1
     If F > 0 Then PosLinea = (PictMatriz.Height - PictMatriz.Top) / F
     If C > 0 Then PosColumna = PictMatriz.width / C
PCol = PictMatriz.Left
For I = 0 To C - 1
    PFil = PictMatriz.Top
    For J = 0 To F - 1
        PictMatriz.Line (PCol, PFil)-(PCol + PosColumna, PFil + PosLinea), Blanco, BF
        PictMatriz.Line (PCol + 0.02, PFil + 0.02)-(PCol + PosColumna - 0.02, PFil + PosLinea - 0.02), Negro, BF
        PictMatriz.Line (PCol + 0.03, PFil + 0.03)-(PCol + PosColumna - 0.03, PFil + PosLinea - 0.03), GMatriz(I, J).color, BF
        Distancia = PosLinea - 0.2
        If Distancia > 0.6 Then Distancia = 0.6
        PictMatriz.PaintPicture LoadPicture(RutaSistema & "\ICONOS\" & GMatriz(I, J).Grafico), PCol + 0.2, PFil + 0.1, Distancia, Distancia
        PictMatriz.FontSize = 6
        PictMatriz.FontBold = True
        PictMatriz.ForeColor = Blanco
        PictMatriz.CurrentX = PCol + Distancia + 0.25
        PictMatriz.CurrentY = PFil + 0.33
        PictMatriz.Print CStr(GMatriz(I, J).Texto)
        PictMatriz.ForeColor = Negro
        PictMatriz.CurrentX = PCol + Distancia + 0.23
        PictMatriz.CurrentY = PFil + 0.3
        PictMatriz.Print CStr(GMatriz(I, J).Texto)
        PFil = PFil + PosLinea
    Next J
    PCol = PCol + PosColumna
Next I

RatonNormal
End Sub

Private Sub PictMatriz_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbLeftButton Then
Si_No = True
PCol = PictMatriz.Left
I = 0
While (I <= C - 1) And Si_No
    PFil = PictMatriz.Top
    J = 0
    While (J <= F - 1) And Si_No
        If ((PCol + 0.02) <= X) And (X <= (PCol + PosColumna - 0.02)) And _
           ((PFil + 0.02) <= y) And (y <= (PFil + PosLinea - 0.02)) Then
            Si_No = False
        End If
        PFil = PFil + PosLinea
        J = J + 1
    Wend
    PCol = PCol + PosColumna
    I = I + 1
Wend
FGraficoMatriz.Caption = X & " x " & y & " - " & I & "," & J
End If
End Sub
