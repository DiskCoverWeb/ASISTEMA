VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ActualizarUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizacion de Usuarios"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstTablas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   105
      TabIndex        =   5
      Top             =   945
      Width           =   3165
   End
   Begin MSAdodcLib.Adodc AdoClave 
      Height          =   330
      Left            =   105
      Top             =   1050
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Clave"
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
      Height          =   855
      Left            =   3360
      Picture         =   "ActuClav.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1050
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Actualizar"
      Enabled         =   0   'False
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
      Left            =   3360
      Picture         =   "ActuClav.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1065
   End
   Begin VB.TextBox TextClaveOld 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "0000000000"
      Top             =   630
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " INGRESAR C.I."
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
      TabIndex        =   4
      Top             =   630
      Width           =   1590
   End
   Begin VB.Label LabelUsuario 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   540
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3165
   End
End
Attribute VB_Name = "ActualizarUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IJ As Long
Dim IdTime As Long
Dim strCnn As String
' Consultamos las cuentas de la tabla
  RatonReloj
  CodigoP = TrimStrg(TextClaveOld)
  LstTablas.Clear
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
        LstTablas.AddItem RstSchema!TABLE_NAME
     End If
     RstSchema.MoveNext
  Loop
  For I = 0 To LstTablas.ListCount - 1
      Si_No = False
      RutaOrigen = LstTablas.List(I)
      sSQL = "SELECT * " _
           & "FROM " & RutaOrigen & " "
      Select_Adodc AdoClave, sSQL
      With AdoClave.Recordset
       For J = 0 To .fields.Count - 1
        If .fields(J).Name = "CodigoU" Then Si_No = True
       Next J
      End With
      If Si_No Then Actualizar_Usuario_Tabla RutaOrigen, CodigoP
  Next I
  Actualizar_Usuario_Tabla "Accesos", CodigoP, True
  Actualizar_Usuario_Tabla "Acceso_Empresa", CodigoP, True
  RatonNormal
  MsgBox "Proceso Exitoso, vuelva a ingresar al programa"
  End
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Activate()
  LabelUsuario.Caption = NombreUsuario
  TextClaveOld = CodigoUsuario
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm ActualizarUsuarios
  ConectarAdodc AdoClave
End Sub

Private Sub TextClaveOld_GotFocus()
  MarcarTexto TextClaveOld
End Sub

Private Sub TextClaveOld_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextClaveOld_LostFocus()
  TextoValido TextClaveOld
  DigVerif = Digito_Verificador(TextClaveOld)
  If Tipo_RUC_CI.Tipo_Beneficiario = "C" Then
     Command1.Enabled = True
     Command1.SetFocus
  Else
     MsgBox "NUMERO DE CEDULA INCORRECTO," & vbCrLf & vbCrLf _
          & "VUELVA HA INGRESAR SIN GUIONES"
     Command1.Enabled = False
     If CodigoPais <> "593" Then
        Command1.Enabled = True
     Else
        TextClaveOld.SetFocus
     End If
  End If
End Sub

Public Sub Actualizar_Usuario_Tabla(NombreTabla, CodAct As String, Optional EsAccesos As Boolean)
  RatonReloj
  If EsAccesos Then
     sSQL = "UPDATE " & NombreTabla & " " _
          & "SET Codigo = '" & CodAct & "' " _
          & "WHERE Codigo = '" & CodigoUsuario & "' "
  Else
     sSQL = "UPDATE " & NombreTabla & " " _
          & "SET CodigoU = '" & CodAct & "' " _
          & "WHERE CodigoU = '" & CodigoUsuario & "' "
  End If
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub
