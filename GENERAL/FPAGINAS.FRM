VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8190
   Begin VB.VScrollBar VScroll1 
      Height          =   3480
      Left            =   7980
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   3465
      Width           =   7995
   End
   Begin VB.OLE OLE1 
      Class           =   "Word.Document.8"
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjWordBasic

Private Sub Archivo_Click()
End Sub

Private Sub Form_Load()
  OLE1.SizeMode = vbOLESizeAutoSize
  OLE1.CreateEmbed "C:\SISTEMA\REPORTES\HOLA.DOC"
End Sub

Private Sub Form_Resize()
Static bFlag As Boolean
If bFlag Then
   OLE1_Resize OLE1.Height, OLE1.Width
Else
   bFlag = True
End If
AdjustScrollBars Me
End Sub

Private Sub HScroll1_Change()
  OLE1.Left = 0 - HScroll1.Value
End Sub

Private Sub OLE1_Resize(HeightNew As Single, WidthNew As Single)
   If HeightNew > Form1.Height Then
      VScroll1.Visible = True
      VScroll1.Max = HeightNew
      VScroll1.LargeChange = HeightNew / (HeightNew / OLE1.Height)
      VScroll1.SmallChange = VScroll1.LargeChange / 10
   Else
      VScroll1.Visible = False
   End If
   
   If WidthNew > Form1.Width Then
      HScroll1.Visible = True
      HScroll1.Max = WidthNew
      HScroll1.LargeChange = WidthNew / (WidthNew / OLE1.Width)
      HScroll1.SmallChange = HScroll1.LargeChange / 10
   Else
      HScroll1.Visible = False
   End If
End Sub

Private Sub VScroll1_Change()
  OLE1.Top = 0 - VScroll1.Value
End Sub

