VERSION 5.00
Begin VB.Form LogInKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE LEGALIZACION Y ACTIVACION"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   5580
      Begin VB.CommandButton CommandButton1 
         Caption         =   "Salir"
         Height          =   225
         Left            =   4830
         TabIndex        =   3
         Top             =   2625
         Width           =   645
      End
      Begin VB.Label LblWeb 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   2
         Top             =   1680
         Width           =   5370
      End
      Begin VB.Label LblMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   5370
      End
   End
End
Attribute VB_Name = "LogInKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
 End
End Sub

Private Sub Form_Load()
  RatonReloj
  Ver_Grafico_Form LogInKey, RutaSistema & "\FORMATOS\INICIO.JPG"
  FrameMsg.Left = (Screen.width - FrameMsg.width) / 2
  FrameMsg.Top = (Screen.Height - FrameMsg.Height) / 3
   
  LblMsg.Caption = "USTED NO ESTA LEGALIZADO LLAME A SU" & vbCrLf _
                 & "PROVEEDOR O A NUESTRAS OFICINAS A" & vbCrLf _
                 & "LOS TELEFONOS: 02-321-0051/099-965-4196/098-910-5300" & vbCrLf _
                 & "EN QUITO - ECUADOR O LOS EMAILS:" & vbCrLf _
                 & "asistencia@diskcoversystem.com / prisma_net@hotmail.es" & vbCrLf _
                 & "PARA SU LEGALIZACION"
  LblWeb.Caption = "www.diskcoversystem.com"
  RatonNormal
End Sub

