VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Cierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Periodo"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataCierre 
      Caption         =   "Cierre"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Picture         =   "Cierre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxFecha 
      Height          =   330
      Left            =   210
      TabIndex        =   1
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha"
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
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1380
   End
End
Attribute VB_Name = "Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  sSQL = "SELECT * FROM Cierre_Mes "
  SelectAdodc DataCierre, sSQL
  With DataCierre.Recordset
      .Edit
      .Fields("Fecha") = MBoxFecha.Text
      .Update
  End With
  FechaCierre = MBoxFecha.Text
  Unload Cierre
End Sub

Private Sub Form_Load()
  CentrarForm Cierre
  DataCierre.DatabaseName = RutaSistema & "\SETEOS.MDB"
  MBoxFecha.Text = FechaCierre
End Sub

Private Sub MBoxFecha_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then Unload Cierre
End Sub

Private Sub MBoxFecha_LostFocus()
  FechaValida MBoxFecha
End Sub
