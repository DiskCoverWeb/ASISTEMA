VERSION 5.00
Begin VB.Form FUnidad 
   BackColor       =   &H000000C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UNIDAD"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6630
   Icon            =   "FUnidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   105
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H008080FF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5565
      Picture         =   "FUnidad.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   " &Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5565
      Picture         =   "FUnidad.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin VB.DriveListBox Drive2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   5370
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " UNIDAD &DESTINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   5370
   End
End
Attribute VB_Name = "FUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constantes
Const Kb As Double = 1024
Const Mb As Double = 1024 * Kb
Const Gb As Double = 1024 * Mb
Const TB As Double = 1024 * Gb

Private Sub Command8_Click()
  FUnidad.Hide
  RutaDestino = UCaseStrg(MidStrg(Drive2.Drive, 1, 2))
  RutaSubDirTemp = RutaDestino
  If ConSubDir Then RutaSubDirTemp = UCaseStrg(Dir1.Path)
  RatonNormal
  Unload FUnidad
End Sub

Private Sub Command8_GotFocus()
  RutaDestino = "C:"
  RutaSubDirTemp = ""
End Sub

Private Sub Command9_Click()
  RutaDestino = "C:"
  RutaSubDirTemp = ""
  Unload FUnidad
End Sub

Private Sub Drive2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Drive2_LostFocus()
  Dir1.Path = UCaseStrg(Drive2.Drive)
End Sub

Private Sub Form_Activate()
  RutaDestino = ""
  If ConSubDir Then Dir1.Visible = True
  Unidad_Temp = UCaseStrg(MidStrg(Drive2.Drive, 1, 2))
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FUnidad
  Me.Caption = " Wmi - Unidades de red "
End Sub

