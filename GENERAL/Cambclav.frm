VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CambClav 
   BorderStyle     =   0  'None
   Caption         =   "2.1.01.01"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   7755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Cambclav.frx":0000
   ScaleHeight     =   4215
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextEmail 
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
      Left            =   1050
      MaxLength       =   60
      TabIndex        =   6
      Top             =   1995
      Width           =   5265
   End
   Begin VB.TextBox TxtPapelImpresora2 
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
      MaxLength       =   60
      TabIndex        =   10
      Top             =   3675
      Width           =   6105
   End
   Begin VB.TextBox TxtImpresora2 
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
      MaxLength       =   60
      TabIndex        =   8
      Top             =   2835
      Width           =   6105
   End
   Begin MSAdodcLib.Adodc AdoClave 
      Height          =   330
      Left            =   210
      Top             =   4410
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
      BackColor       =   &H00C0E0FF&
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
      Left            =   6510
      Picture         =   "Cambclav.frx":9254
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1260
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Grabar"
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
      Left            =   6510
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cambclav.frx":9B1E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox TextClaveNew 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2415
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "***"
      Top             =   1470
      Width           =   3480
   End
   Begin VB.TextBox TextClaveOld 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2415
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "***"
      Top             =   945
      Width           =   3480
   End
   Begin MSAdodcLib.Adodc AdoMySQLClave 
      Height          =   330
      Left            =   2100
      Top             =   4410
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
      Caption         =   "MySQLClave"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " EMAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   210
      TabIndex        =   5
      Top             =   1995
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TIPO DE PAPEL DE LA IMPRESORA SECUNDARIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   210
      TabIndex        =   9
      Top             =   3360
      Width           =   6105
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DE LA IMPRESORA SECUNDARIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   2520
      Width           =   6105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " CLAVE ANTERIOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   210
      TabIndex        =   1
      Top             =   945
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " CLAVE NUEVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Left            =   210
      TabIndex        =   3
      Top             =   1470
      Width           =   2115
   End
   Begin VB.Label LabelUsuario 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   540
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   6105
   End
End
Attribute VB_Name = "CambClav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vUsuario As String
Dim vClave As String
Dim vListaEmpresas As String
Dim vMensaje As String

Private Sub Command1_Click()
Dim AdoLstEmp As ADODB.Recordset
Dim MsgAux As String
Dim Contactos As String

RatonReloj
TextoValido TxtImpresora2
TextoValido TxtPapelImpresora2
TextoValido TextEmail

If AdoClave.Recordset.RecordCount > 0 Then
   AdoClave.Recordset.fields("Clave") = TextClaveNew
   AdoClave.Recordset.fields("EmailUsuario") = TextEmail
   AdoClave.Recordset.fields("Impresora_Defecto_2") = TxtImpresora2
   AdoClave.Recordset.fields("Papel_Impresora_2") = TxtPapelImpresora2
   AdoClave.Recordset.Update
   
   If AdoMySQLClave.Recordset.RecordCount > 0 Then
      AdoMySQLClave.Recordset.fields("Clave") = TextClaveNew
      AdoMySQLClave.Recordset.fields("Email") = TextEmail
      AdoMySQLClave.Recordset.Update
   End If
   
   Contactos = ""
   Insertar_Cadena Contactos, Telefono1
   Insertar_Cadena Contactos, Telefono2
   If Len(RazonSocial) > 1 Then MsgAux = RazonSocial Else MsgAux = Empresa
   vMensaje = MensajeAutomatizado
   vMensaje = Replace(vMensaje, "Nombre_Usuario", NombreUsuario)
   vMensaje = Replace(vMensaje, "Mensaje_Comunicado", "")
   vMensaje = Replace(vMensaje, "Representante_Legal", NombreGerente)
   vMensaje = Replace(vMensaje, "Numero_Telefono", Contactos)
   vMensaje = Replace(vMensaje, "Emails", EmailProcesos)
   vMensaje = Replace(vMensaje, "Razon_Social", MsgAux)
   vMensaje = Replace(vMensaje, vbCrLf, "<br>")
   
   vListaEmpresas = ""
   sSQL = "SELECT E.Empresa, M.Aplicacion " _
        & "FROM Empresas As E " _
        & "INNER JOIN Acceso_Empresa As AE " _
        & "ON E.Item = AE.Item  " _
        & "INNER JOIN Modulos As M  " _
        & "ON AE.Modulo = M.Modulo  " _
        & "WHERE AE.Codigo = '" & CodigoUsuario & "'  " _
        & "ORDER BY E.Empresa, M.Aplicacion "
   Select_AdoDB AdoLstEmp, sSQL
   If AdoLstEmp.RecordCount > 0 Then
      Do While Not AdoLstEmp.EOF
         vListaEmpresas = vListaEmpresas & "<tr>"
         vListaEmpresas = vListaEmpresas & "<td>" & AdoLstEmp.fields("Empresa") & "</td>"
         vListaEmpresas = vListaEmpresas & "<td>" & AdoLstEmp.fields("Aplicacion") & "</td>"
         vListaEmpresas = vListaEmpresas & "</tr>"
         AdoLstEmp.MoveNext
      Loop
   End If
   
  'Datos del destinatario de mails
   TMail.Mensaje = NombreComercial & vbCrLf _
                 & RazonSocial & vbCrLf _
                 & Telefono1 & "/" & Telefono1 & vbCrLf _
                 & "Dir. " & Direccion & vbCrLf _
                 & UCaseStrg(NombreCiudad) & "-" & UCaseStrg(NombrePais) & vbCrLf
   TMail.MensajeHTML = Leer_Archivo_Texto(RutaSistema & "\FORMATOS\credenciales.html")
   TMail.MensajeHTML = Replace(TMail.MensajeHTML, "vEntidad", RUC)
   TMail.MensajeHTML = Replace(TMail.MensajeHTML, "vUsuario", vUsuario)
   TMail.MensajeHTML = Replace(TMail.MensajeHTML, "vClave", TextClaveNew)
   TMail.MensajeHTML = Replace(TMail.MensajeHTML, "vEntidad", NombreEntidad)
   TMail.MensajeHTML = Replace(TMail.MensajeHTML, "vListaEmpresas", vListaEmpresas)
   TMail.MensajeHTML = TMail.MensajeHTML & "<div>" & vMensaje & "</div>"
  'Enviamos lista de mails
   TMail.TipoDeEnvio = Ninguno
   TMail.para = ""
   TMail.de = ""
   TMail.Usuario = ""
   TMail.Password = ""
   TMail.ListaMail = 6
   TMail.Asunto = "Envio Credenciales ingresos al sistema"
   Insertar_Mail TMail.para, TextEmail
   Insertar_Mail TMail.para, Lista_De_Correos(TMail.ListaMail).Correo_Electronico
   TMail.Adjunto = ""
   FEnviarCorreos.Show vbModal
   RatonNormal
   MsgBox UCaseStrg(NombreUsuario & ", su requerimiento se ha procesado con exito")
End If
RatonNormal
Unload CambClav
End Sub

Private Sub Command2_Click()
  Unload CambClav
End Sub

Private Sub Form_Activate()
   'Leemos datos del Usuario en SQLSERVER
    IDEUsuario = Ninguno
    TextClaveOld = "**********"
    TextClaveNew = ""
    LabelUsuario.Caption = "Usuario: " & vbCrLf & NombreUsuario
    sSQL = "SELECT Codigo, Usuario, Clave, EmailUsuario, Impresora_Defecto_2, Papel_Impresora_2, ID " _
         & "FROM Accesos " _
         & "WHERE UCaseStrg(Nombre_Completo) = '" & UCaseStrg(NombreUsuario) & "' "
    Select_Adodc AdoClave, sSQL
    If AdoClave.Recordset.RecordCount > 0 Then
       IDEUsuario = AdoClave.Recordset.fields("Codigo")
       vClave = AdoClave.Recordset.fields("Clave")
       vUsuario = AdoClave.Recordset.fields("Usuario")
       TextEmail = AdoClave.Recordset.fields("EmailUsuario")
       TxtImpresora2 = AdoClave.Recordset.fields("Impresora_Defecto_2")
       TxtPapelImpresora2 = AdoClave.Recordset.fields("Papel_Impresora_2")
    End If
  
   'Leemos datos del Usuario en MySQL
    sSQL = "SELECT CI_NIC, Clave, Email, ID " _
         & "FROM acceso_usuarios " _
         & "WHERE CI_NIC = '" & IDEUsuario & "' "
    Select_Adodc AdoMySQLClave, sSQL
    RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm CambClav
  Redondear_Formulario CambClav, 35
  ConectarAdodc AdoClave
  ConectarAdodc_MySQL AdoMySQLClave
End Sub

Private Sub TextClaveNew_GotFocus()
  MarcarTexto TextClaveNew
End Sub

Private Sub TextClaveNew_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextClaveOld_GotFocus()
  MarcarTexto TextClaveOld
End Sub

Private Sub TextClaveOld_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextClaveOld_LostFocus()
  If TextClaveOld <> vClave Then
     MsgBox "Clave incorrecta"
     Command1.Enabled = False
  Else
     Command1.Enabled = True
  End If
End Sub

Private Sub TextEmail_GotFocus()
   MarcarTexto TextEmail
End Sub

Private Sub TextEmail_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextEmail_LostFocus()
   TextoValido TextEmail
End Sub

Private Sub TxtImpresora2_GotFocus()
  MarcarTexto TxtImpresora2
End Sub

Private Sub TxtImpresora2_LostFocus()
  TextoValido TxtImpresora2
End Sub

Private Sub TxtPapelImpresora2_GotFocus()
  MarcarTexto TxtPapelImpresora2
End Sub

Private Sub TxtPapelImpresora2_LostFocus()
  TextoValido TxtPapelImpresora2
End Sub
