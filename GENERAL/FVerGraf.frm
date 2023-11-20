VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "Comctl32.ocx"
Begin VB.Form FVerGrafico 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   DrawStyle       =   5  'Transparent
   Icon            =   "FVerGraf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1164
      ButtonWidth     =   2778
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Vertical"
            Key             =   "Vertical"
            Object.ToolTipText     =   "Imprime Verticalmente"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Horizontal"
            Key             =   "Horizontal"
            Object.ToolTipText     =   "Imprime Horizontalmente"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Medidas"
            Key             =   "Medidas"
            Object.ToolTipText     =   "Medidas del Gráfico"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pagina &Anterior"
            Key             =   "Anterior"
            Object.ToolTipText     =   "Pagina Anterior"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pagina &Siguiente"
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Pagina Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Quit"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   228
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   5370
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6030
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   228
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   5  'Transparent
      Height          =   6000
      Left            =   210
      ScaleHeight     =   5940
      ScaleWidth      =   11295
      TabIndex        =   2
      Top             =   630
      Width           =   11355
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawStyle       =   5  'Transparent
         Height          =   5685
         Left            =   0
         ScaleHeight     =   5625
         ScaleWidth      =   11085
         TabIndex        =   4
         Top             =   0
         Width           =   11145
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerGraf.frx":128C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FVerGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RutaDibujoVerGrafico As String
Dim NombFilePict As String
Dim NONE As Boolean

Private Sub Form_Activate()
'Dibujar un Grafico y hacer scroll
 NONE = False
 Pagina = 1
 NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".GIF"
 RutaOrigen = RutaSistema & "\PRINTER\" & NombFilePict
 RutaOrigen = UCase$(RutaOrigen)
 RutaDibujoVerGrafico = RutaOrigen
 If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
 Const PIXEL = 7
 FVerGrafico.Caption = RutaOrigen
 'Establece las propiedades incluidas aqui por simplicidad
 FVerGrafico.ScaleMode = PIXEL
 Picture1.ScaleMode = PIXEL
 Picture2.ScaleMode = PIXEL
 'AutoSize se pone a TRUE de forma que los límites de Picture2 son expandidos
 'de hasta el tamaño del BitMap
  Picture2.AutoSize = True
 'Liberarse de los bordes que molestan
 Picture1.BorderStyle = NONE
 Picture2.BorderStyle = NONE
 'Cargas el Dibujo que quieres visualizar
 Picture2.FontBold = True
 Picture2.Picture = LoadPicture(RutaOrigen)
 ''' Rotacion_Texto Picture2, 15, 15, 30, 12, 14, "30 Grados Hola como estas", TipoTimes
 'inicializa la posicion de ambos dibujos
 Picture1.Move 1, 1, ScaleWidth - VScroll1.width, ScaleHeight - HScroll1.Height
 Picture2.Move 1, 1
 'posicion de la barra horizontal de desplazamiento
 HScroll1.Top = Picture1.Height
 HScroll1.Left = VScroll1.width
 HScroll1.width = Picture1.width
 'posicion la barra vertical desplazamiento
 VScroll1.Top = Toolbar1.Height
 VScroll1.Left = 0 ' Picture1.Width
 VScroll1.Height = ScaleHeight - (HScroll1.Height + 1)
 'VScroll1.Width = Picture1.Height
 'Establece el maximo valor para las barra desplazamiento
 HScroll1.Max = Picture2.width - Picture1.width
 VScroll1.Max = Picture2.Height - Picture1.Height
 'Determina si el dibujo Hijo llenara la pantalla
 'si no lo hace no habra necesidad de barra de desplazamiento
 VScroll1.Enabled = (Picture1.Height < Picture2.Height)
 HScroll1.Enabled = (Picture1.width < Picture2.width)
 Toolbar1.Buttons.Item(3).Caption = Format(Picture2.width, "00.00") & " x " & Format(Picture2.Height, "00.00")
'Fin del Dibujo

 RatonNormal
 Else
   MsgBox "Presentación erronea"
   Unload Me
 End If
End Sub

Private Sub VScroll1_Change()
  'Picture2.Top  es puesto a su valor negativo porque cuando se mueve la barra de
  'desplazamiento hacia abajo, la pantalla de debe desplazar hacia arriba, mostrando
  'mas de la parte inferior de la pantalla y viceversa cuando lo hacemos hacia arriba
  Picture2.Top = -VScroll1.value
  If VScroll1.value = 0 Then Picture2.Top = 1
End Sub

Private Sub HScroll1_Change()
  'Picture2.Height  es puesto a su valor negativo porque cuando se mueve la barra de
  'desplazamiento a la derecha, la pantalla de debe desplazar a la izquierda, mostrando
  'mas de la derecha de la pantalla y viceversa cuando lo hacemos a la izquierda
  Picture2.Left = -HScroll1.value
  If HScroll1.value = 0 Then Picture2.Left = 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim HayMasGraficos As Boolean
  Select Case Button.key
    Case "Horizontal"
         ImpHorizontal
    Case "Vertical"
         ImpVertical
    Case "Anterior"
         Pagina = Pagina - 1
         If Pagina < 1 Then Pagina = 1
         NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".BMP"
         RutaOrigen = UCase$(RutaSistema & "\PRINTER\" & NombFilePict)
         RutaDibujoVerGrafico = RutaOrigen
         FVerGrafico.Caption = RutaOrigen
         If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
            Picture2.Picture = LoadPicture(RutaOrigen)
         Else
            Pagina = 1
            NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".BMP"
            RutaOrigen = UCase$(RutaSistema & "\PRINTER\" & NombFilePict)
            RutaDibujoVerGrafico = RutaOrigen
            If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
               Picture2.Picture = LoadPicture(RutaOrigen)
            End If
         End If
    Case "Siguiente"
         Pagina = Pagina + 1
         NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".BMP"
         RutaOrigen = UCase$(RutaSistema & "\PRINTER\" & NombFilePict)
         RutaDibujoVerGrafico = RutaOrigen
         FVerGrafico.Caption = RutaOrigen
         If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
            Picture2.Picture = LoadPicture(RutaOrigen)
         Else
            Pagina = 1
            NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".BMP"
            RutaOrigen = UCase$(RutaSistema & "\PRINTER\" & NombFilePict)
            RutaDibujoVerGrafico = RutaOrigen
            If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
               Picture2.Picture = LoadPicture(RutaOrigen)
            End If
         End If
    Case "Salir"
         RatonReloj
         HayMasGraficos = True
         Pagina = 1
         Do While HayMasGraficos
            NombFilePict = CodigoUsuario & NumEmpresa & NumModulo & Format(Pagina, "00") & ".BMP"
            RutaOrigen = UCase$(RutaSistema & "\PRINTER\" & NombFilePict)
            RutaDibujoVerGrafico = RutaOrigen
            If Dir(RutaDibujoVerGrafico, vbNormal) <> "" Then
               Kill RutaDibujoVerGrafico
            Else
               HayMasGraficos = False
            End If
            Pagina = Pagina + 1
         Loop
         'If Ucase$(Dir(RutaDibujoVerGrafico, vbNormal)) <> "" Then Kill RutaDibujoVerGrafico
         RatonNormal
         Unload Me
  End Select
End Sub

Public Sub ImpHorizontal()
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION HORIZONTAL"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 0.5: InicioY = 0
   Escala_Centimetro 2, TipoTimes, 8
   Pagina = 1
   AnchoPict = Picture2.width
   AltoPict = Picture2.Height
   If AnchoPict > (Printer.ScaleWidth - 1) Then AnchoPict = Printer.ScaleWidth - 1
   If AltoPict > (Printer.ScaleHeight - 0.5) Then AltoPict = Printer.ScaleHeight - 0.5
   'MsgBox AnchoPict
   PrinterPaint RutaOrigen, 1, 0.5, AnchoPict, AltoPict
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub

Public Sub ImpVertical()
Dim AnchoPict As Single
Dim AltoPict As Single
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION VERTICAL"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 0.5: InicioY = 0
   Escala_Centimetro 1, TipoTimes, 9
   Pagina = 1
   AnchoPict = Picture2.width
   AltoPict = Picture2.Height
   If AnchoPict > Printer.ScaleWidth - 0.5 Then AnchoPict = Printer.ScaleWidth - 0.5
   If AltoPict > Printer.ScaleHeight - 1 Then AltoPict = Printer.ScaleHeight - 1
   PrinterPaint RutaOrigen, 0.2, 0.2, AnchoPict, AltoPict
   RatonNormal
   MensajeEncabData = ""
   Printer.EndDoc
   Exit Sub
Errorhandler:
             RatonNormal
             ErrorDeImpresion
             Exit Sub
Else
   RatonNormal
End If
End Sub
