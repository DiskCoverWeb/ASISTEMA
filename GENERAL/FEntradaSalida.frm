VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FTarjetas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REGISTRO DE ENTRADA/SALIDA"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   1380
      Left            =   105
      Picture         =   "FEntradaSalida.frx":0000
      ScaleHeight     =   1320
      ScaleWidth      =   3525
      TabIndex        =   1
      Top             =   105
      Width           =   3585
   End
   Begin MSAdodcLib.Adodc AdoTarjeta 
      Height          =   330
      Left            =   105
      Top             =   210
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Tarjeta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtTarjeta 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   105
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1575
      Width           =   5055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1410
      Left            =   3675
      Picture         =   "FEntradaSalida.frx":1363
      Stretch         =   -1  'True
      Top             =   105
      Width           =   1440
   End
End
Attribute VB_Name = "FTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  RatonNormal
  sSQL = "SELECT CR.*,C.Cliente " _
       & "FROM Catalogo_Rol_Pagos As CR,Clientes As C " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND CR.Codigo = C.Codigo "
  SelectAdodc AdoTarjeta, sSQL
  TxtTarjeta.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FTarjetas
  ConectarAdodc AdoTarjeta
End Sub

Private Sub Picture1_GotFocus()
 CodigoP = CaracteresValidos(TxtTarjeta)
 'MsgBox CodigoP & vbCrLf & Len(TxtTarjeta) & vbCrLf & TxtTarjeta
 If Len(CodigoP) > 1 Then
   With AdoTarjeta.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Tarjeta = '" & CodigoP & "' ")
        If Not .EOF Then
           MiTiempo = Time
           Cadena = "USUARIO: " & UCase(.Fields("Cliente")) & vbCrLf & vbCrLf _
                  & "HORA DE REGISTRO: " & Format(MiTiempo, FormatoTimes) & vbCrLf & vbCrLf _
                  & "INGRESE SU PROCESO:"
           CodigoB = InputBox(Cadena, "REGISTRO POR TARJETA", "")
           If CodigoB = "" Then CodigoB = Ninguno
           CodigoCli = .Fields("Codigo")
           SetAdoAddNew "Trans_Entrada_Salida"
           SetAdoFields "ES", "H"
           SetAdoFields "Codigo", CodigoCli
           SetAdoFields "Hora", Format(MiTiempo, FormatoTimes)
           SetAdoFields "Fecha", FechaSistema
           SetAdoFields "Proceso", "REGISTRO POR TARJETA"
           SetAdoFields "Tarea", Trim$(Mid$(CodigoB, 1, 50))
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "Periodo", Periodo_Contable
           SetAdoUpdate
           TxtTarjeta.SetFocus
        Else
           MsgBox "Codigo No Asignado"
           TxtTarjeta.SetFocus
        End If
    End If
   End With
 Else
   MsgBox "Tarjeta no Asignada"
   Unload FTarjetas
 End If
End Sub

Private Sub TxtTarjeta_GotFocus()
  TxtTarjeta = ""
End Sub

Private Sub TxtTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
  'PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload FTarjetas
End Sub

