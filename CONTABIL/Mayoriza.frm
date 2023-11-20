VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Mayorizacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mayor Analitico"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Left            =   3360
      Picture         =   "Mayoriza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mayorizar"
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
      Left            =   2310
      Picture         =   "Mayoriza.frx":0282
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   525
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Desde:"
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
      Width           =   750
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hasta:"
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
      TabIndex        =   2
      Top             =   525
      Width           =   750
   End
End
Attribute VB_Name = "Mayorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  FechaInicial = MBoxFechaI
  FechaFinal = MBoxFechaF
  Codigo1 = "1": Codigo2 = "9"
  Unload Mayorizacion
  Mayorizar.Show
End Sub

Private Sub Command2_Click()
  Unload Mayorizacion
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarArribaForm Mayorizacion
End Sub

Private Sub MBoxFechaF_GotFocus()
  MBoxFechaF.Text = LimpiarFechas
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub MBoxFechaI_GotFocus()
  MBoxFechaI.Text = LimpiarFechas
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub
