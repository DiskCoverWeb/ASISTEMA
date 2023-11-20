VERSION 5.00
Begin VB.Form PagPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   8820
      Picture         =   "Pagprint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Comunicarse con WALTER VACA PRIETO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2955
      Left            =   3570
      TabIndex        =   2
      Top             =   945
      Width           =   5160
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos: 02-605-2430/09-9965-4196/09-8910-5300 Emails: diskcover@msn.com - diskcoversystem@msn.com prisma_net@hotmail.es"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2430
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   4950
      End
   End
   Begin VB.PictureBox PictLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   105
      ScaleHeight     =   3.545
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   5.953
      TabIndex        =   4
      Top             =   1050
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Este Programa esta autorizado para ser "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   9780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Este Programa esta autorizado para ser utilizado por:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9780
   End
End
Attribute VB_Name = "PagPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload PagPrint
End Sub

Private Sub Form_Activate()
  Label2.Caption = "''" & Empresa & "''"
  Label3.Caption = "Asistencia Técnica al teléfono PBX: 02-605-2430" & vbCrLf _
                 & "Teléfono Programador:  09-9965-4196/09-8910-5300" & vbCrLf _
                 & "Email Gerencia: prisma_net@hotmail.es/diskcover@msn.com" & vbCrLf _
                 & "Email Asesores: diskcover_contabilidad@outlook.com diskcoversystem@msn.com"
  If LogoTipo <> "" And Dir(LogoTipo) <> "" Then PictLogo.PaintPicture LoadPicture(LogoTipo), 0.01, 0.01, 6, 3
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm PagPrint
   Ver_Grafico_Form PagPrint, RutaSistema & "\FONDOS\DISKCOVS.JPG"
End Sub

