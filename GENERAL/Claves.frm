VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form IngClaves 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUPERVISOR"
   ClientHeight    =   1380
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   3600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Claves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   1170
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3375
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
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
         Height          =   855
         Left            =   2100
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Claves.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1065
      End
      Begin VB.TextBox TextClave 
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
         IMEMode         =   3  'DISABLE
         Left            =   210
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAVE DE ACCESO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   2
         Top             =   315
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoSup 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Sup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "IngClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Intentos As Integer
Dim Claves As String

Private Sub Command1_Click()
  Unload IngClaves
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Accesos " _
       & "WHERE Usuario = '" & TipoSuper & "' "
  Select_Adodc AdoSup, sSQL
  If AdoSup.Recordset.RecordCount > 0 Then
     ClaveGeneral = AdoSup.Recordset.Fields("Clave")
     IngClaves.Caption = AdoSup.Recordset.Fields("Nombre_Completo")
  End If
  Intentos = 0
  ResultClaveSup = False
End Sub

Private Sub Form_Load()
  CentrarForm IngClaves
  Intentos = 0
  ResultClaveSup = False
  ConectarAdodc AdoSup
  RatonNormal
End Sub

Private Sub TextClave_Change()
  If Len(TextClave.Text) >= TextClave.MaxLength Then SendKeys "{TAB}", False
End Sub

Private Sub TextClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload IngClaves
End Sub

Private Sub TextClave_LostFocus()
If TextClave.Text <> "" Then
   Intentos = Intentos + 1
   If (TextClave.Text = ClaveGeneral) And (Intentos < 3) Then
      ResultClaveSup = True
      Unload IngClaves
   ElseIf Intentos >= 3 Then
      Cadena = "Sr(a). " & NombreUsuario & ": " & vbCrLf _
             & Space(10) & "Usted no está autorizado" & vbCrLf _
             & Space(10) & "a ingresar al esta opción." & vbCrLf
      MsgBox Cadena
      Unload IngClaves
   Else
      Cadena = "Sr(a). " & NombreUsuario & ": " & vbCrLf _
             & Space(10) & "Clave incorrecta," & vbCrLf
      MsgBox Cadena
      Claves = ""
      TextClave.Text = ""
      TextClave.SetFocus
   End If
End If
End Sub

