VERSION 5.00
Begin VB.Form DatosEmp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   105
      ScaleHeight     =   2220
      ScaleWidth      =   3270
      TabIndex        =   4
      Top             =   105
      Width           =   3270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SALIR"
      Height          =   540
      Left            =   2625
      TabIndex        =   2
      Top             =   4725
      Width           =   1800
   End
   Begin VB.PictureBox PictureLogo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   3570
      ScaleHeight     =   2220
      ScaleWidth      =   3270
      TabIndex        =   0
      Top             =   105
      Width           =   3270
   End
   Begin VB.Label Label1 
      Height          =   750
      Left            =   105
      TabIndex        =   3
      Top             =   3885
      Width           =   6945
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   6945
   End
End
Attribute VB_Name = "DatosEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload DatosEmp
End Sub

Private Sub Form_Load()
   CentrarForm DatosEmp
   Cadena = "Para cualquier consulta llamar a las Empresas:" & Chr(13) & Chr(13)
   Cadena = Cadena & "DiskCover" & Chr(13) & Chr(13)
   Cadena = Cadena & "Teléf: 02-528119/09-823551"
   Cadena = Cadena & "Quito - Ecuador"
   Label4.Caption = Cadena
   'PictureLogo.Picture = LoadPicture(LogoTipo)
   'Picture1.Picture = LoadPicture(RutaSistema & "\FORMATOS\DISKCOVE.WMF")
End Sub

