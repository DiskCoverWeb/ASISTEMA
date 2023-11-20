VERSION 5.00
Begin VB.Form FMensaje 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2400
   ClientLeft      =   30
   ClientTop       =   45
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "PRESIONE ENTER O CLICK PARA CONTINUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      MaskColor       =   &H80000018&
      TabIndex        =   0
      Top             =   1890
      Width           =   7050
   End
   Begin VB.PictureBox Picture1 
      Height          =   1800
      Left            =   0
      Picture         =   "FMensaje.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   2
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label LMensaje 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1800
      Left            =   1785
      TabIndex        =   1
      Top             =   0
      Width           =   5475
   End
End
Attribute VB_Name = "FMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   FMensaje.Hide
End Sub

Private Sub Form_Activate()
   LMensaje.Caption = Mensaje
   Command1.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FMensaje
End Sub

Private Sub LMensaje_Click()
   FMensaje.Hide
End Sub

Private Sub Picture1_Click()
   FMensaje.Hide
End Sub
