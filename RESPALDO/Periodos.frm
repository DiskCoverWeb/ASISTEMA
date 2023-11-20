VERSION 5.00
Begin VB.Form Periodos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAMBIO DE PERIODOS"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4530
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3360
      Picture         =   "Periodos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1155
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambio de &Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   3360
      Picture         =   "Periodos.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   1065
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   105
      TabIndex        =   5
      Top             =   420
      Width           =   3165
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA EL PERIODO"
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
      Top             =   105
      Width           =   3165
   End
   Begin VB.Label LabelTotSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   9555
      TabIndex        =   4
      Top             =   6930
      Width           =   2010
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Debe - Haber:"
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
      Left            =   7665
      TabIndex        =   3
      Top             =   6930
      Width           =   1800
   End
End
Attribute VB_Name = "Periodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  RatonReloj
  RutaEmpresa = Dir1.Path
  RatonNormal
  Unload Periodos
End Sub

Private Sub Command3_Click()
  Unload Periodos
End Sub

Private Sub Form_Activate()
  Dir1.Path = RutaEmpresa
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Periodos
End Sub

