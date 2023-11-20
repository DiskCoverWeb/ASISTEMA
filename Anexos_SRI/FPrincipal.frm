VERSION 5.00
Begin VB.Form FPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexos Transaccionales"
   ClientHeight    =   1485
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   336
      Left            =   2376
      TabIndex        =   3
      Top             =   864
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   336
      Left            =   2376
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   324
      Width           =   1092
   End
   Begin VB.Frame FrmOpciones 
      Caption         =   "Seleccione Opción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1308
      Left            =   216
      TabIndex        =   0
      Top             =   108
      Width           =   1956
      Begin VB.OptionButton OpcVentas 
         Caption         =   "Ventas"
         Height          =   228
         Left            =   216
         TabIndex        =   4
         Top             =   756
         Width           =   984
      End
      Begin VB.PictureBox OpcCompras 
         Height          =   336
         Left            =   216
         ScaleHeight     =   270
         ScaleWidth      =   1035
         TabIndex        =   1
         Top             =   324
         Width           =   1092
      End
   End
End
Attribute VB_Name = "FPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Opc As String
Private Sub Command1_Click()
    FCompras.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()

End Sub

Private Sub OpcCompras_click()
    Opc = "Compras"
    MsgBox Opc
End Sub

Private Sub OpcVentas_click()
    Opc = "Ventas"
    MsgBox Opc
End Sub
