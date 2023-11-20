VERSION 5.00
Begin VB.Form No_LPTs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Impresora por defecto"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   105
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1365
      Width           =   3585
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Prueba"
      Height          =   435
      Left            =   3990
      TabIndex        =   6
      Top             =   1365
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   435
      Left            =   3990
      TabIndex        =   5
      Top             =   2625
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3990
      TabIndex        =   4
      Top             =   1995
      Width           =   1380
   End
   Begin VB.ListBox LstLPTs 
      Height          =   1035
      Left            =   105
      TabIndex        =   0
      Top             =   2100
      Width           =   3585
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto de Prueba:"
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
      TabIndex        =   7
      Top             =   1050
      Width           =   3585
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Seleccione la impresora por Defecto"
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
      TabIndex        =   3
      Top             =   1785
      Width           =   3585
   End
   Begin VB.Label LabelLPTs 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   5475
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Impresora por Defecto"
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
      Top             =   105
      Width           =   5475
   End
End
Attribute VB_Name = "No_LPTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
For Each Impresora In Printers
   If Impresora.DeviceName = LstLPTs.Text Then
      Set Printer = Impresora
      Exit For
   End If
Next
Cadena = "Puerto de Comunicacion: " & Printer.Port & Chr(13)
Cadena = Cadena & "                        Impresora: " & Printer.DeviceName
LabelLPTs.Caption = Cadena
End Sub

Private Sub Command2_Click()
    Unload No_LPTs
End Sub

Private Sub Command3_Click()
On Error GoTo Errorhandler
'Establecemos Espacios y seteos de impresion
EscalaCentimetro 1, TipoTimes, 10
PosLinea = 0
For I = 6 To 30
   Printer.FontSize = I
   Printer.CurrentX = 2: Printer.CurrentY = PosLinea
   Printer.Print Text1.Text
   PosLinea = PosLinea + 1
Next I
Printer.EndDoc
MsgBox "Se ha impreso correctamente la prueba en: " & Printer.DeviceName
Exit Sub
Errorhandler:
    MsgBox "Error: No pudo imprimir correctamente en " & Printer.DeviceName
    Exit Sub
End Sub

Private Sub Form_Activate()
ReDim Lpts(Printers.Count) As Printer
Cadena = "Puerto de Comunicacion: " & Printer.Port & Chr(13)
Cadena = Cadena & "                        Impresora: " & Printer.DeviceName
LabelLPTs.Caption = Cadena
LstLPTs.Clear: I = 0
For Each Lpts(I) In Printers
    LstLPTs.AddItem Lpts(I).DeviceName
    I = I + 1
Next
End Sub

Private Sub Form_Load()
   CentrarForm No_LPTs
   Set Impresora = Printer
End Sub
