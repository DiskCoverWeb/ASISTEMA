VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FBenef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATOS PRINCIPALES DEL BENEFICIARIO"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.Data DataBenef 
      Caption         =   "Benef"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1260
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox TextCiudad 
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
      Left            =   1155
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1050
      Width           =   1485
   End
   Begin VB.TextBox TextDirecc 
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
      Left            =   3990
      MaxLength       =   60
      TabIndex        =   15
      Top             =   1050
      Width           =   6315
   End
   Begin MSMask.MaskEdBox MBoxRUC 
      Height          =   330
      Left            =   1470
      TabIndex        =   5
      Top             =   630
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#########-#-###"
      Mask            =   "#########-#-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxTelefono 
      Height          =   330
      Left            =   4620
      TabIndex        =   7
      Top             =   630
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   327680
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
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxCelular 
      Height          =   330
      Left            =   7035
      TabIndex        =   9
      Top             =   630
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   327680
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
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBoxFAX 
      Height          =   330
      Left            =   9030
      TabIndex        =   11
      Top             =   630
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   327680
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
      Format          =   "##-###-###"
      Mask            =   "##-###-###"
      PromptChar      =   "0"
   End
   Begin VB.Label LabelCodigo 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   9030
      TabIndex        =   3
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CODIGO:"
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
      Left            =   8085
      TabIndex        =   2
      Top             =   210
      Width           =   960
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION:"
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
      Left            =   2730
      TabIndex        =   14
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FAX:"
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
      Left            =   8400
      TabIndex        =   10
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CELULAR:"
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
      Left            =   5985
      TabIndex        =   8
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TELEFONO:"
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
      Left            =   3465
      TabIndex        =   6
      Top             =   630
      Width           =   1170
   End
   Begin VB.Label LabelBenef 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   6315
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BENEFICIARIO:"
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
      Width           =   1485
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " R.U.C. / C.I."
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
      TabIndex        =   4
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CIUDAD:"
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
      TabIndex        =   12
      Top             =   1050
      Width           =   960
   End
End
Attribute VB_Name = "FBenef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim PosIzq As Single
   PosIzq = ((Screen.Width - FBenef.Width) / 2)
   If PosIzq < 0 Then PosIzq = 0
   FBenef.Left = PosIzq: FBenef.Top = 1100
   LabelBenef.Caption = Beneficiario
   DataBenef.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub MBoxCelular_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxFAX_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxRUC_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCiudad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCiudad_LostFocus()
  TextoValido TextCiudad, False
End Sub

Private Sub TextDirecc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextDirecc_LostFocus()
  TextoValido TextDirecc, False
  Mensajes = "Está seguro de haber ingresado los datos"
  Titulo = "FORMULARIO DE GRABACION"
  If BoxMensaje = 6 Then
     Numero = ReadSetDataNum("Beneficiarios", True, True)
     CodigoBenef = FormatoCodigo(LabelBenef.Caption, Numero)
     SelectData DataBenef, "Beneficiarios", False
     With DataBenef.Recordset
         .AddNew
         .Fields("TC") = "N"
         .Fields("Codigo") = CodigoBenef
         .Fields("Beneficiario") = Mid(LabelBenef.Caption, 1, 30)
         .Fields("Ciudad") = TextCiudad.Text
         .Fields("Direccion") = TextDirecc.Text
         .Fields("RUC_CI") = MBoxRUC.Text
         .Fields("Telefono") = MBoxTelefono.Text
         .Fields("Celular") = MBoxCelular.Text
         .Fields("FAX") = MBoxFAX.Text
         .Update
     End With
     Unload FBenef
  End If
End Sub
