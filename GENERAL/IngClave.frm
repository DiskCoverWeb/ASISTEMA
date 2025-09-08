VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form IngClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CREACION DE USUARIOS"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "IngClave.frx":0000
   ScaleHeight     =   5580
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextEmail 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      MaxLength       =   60
      TabIndex        =   3
      Top             =   1365
      Width           =   7785
   End
   Begin VB.Frame FrmEstEmi 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   11
      Top             =   3045
      Visible         =   0   'False
      Width           =   7785
      Begin VB.TextBox TxtEmailEstEmi 
         Height          =   330
         Left            =   1155
         MaxLength       =   60
         TabIndex        =   16
         Top             =   630
         Width           =   6420
      End
      Begin VB.TextBox TxtDireccionEstab 
         Height          =   330
         Left            =   210
         MaxLength       =   60
         TabIndex        =   18
         Top             =   1365
         Width           =   7365
      End
      Begin VB.TextBox TxtLogoTipoEstab 
         Height          =   330
         Left            =   5460
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1785
         Width           =   2115
      End
      Begin VB.TextBox TxtTelefonoEstab 
         Height          =   330
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1785
         Width           =   2115
      End
      Begin VB.TextBox TxtNumSerieUno 
         Height          =   336
         Left            =   210
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
         Top             =   630
         Width           =   435
      End
      Begin VB.TextBox TxtNumSerieDos 
         Height          =   336
         Left            =   630
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "001"
         ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label24 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CORREO ELECTRONICO"
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   1155
         TabIndex        =   15
         Top             =   315
         Width           =   6420
      End
      Begin VB.Label Label25 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DIRECCION"
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   1050
         Width           =   7365
      End
      Begin VB.Label Label26 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO:"
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   19
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label Label23 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " LOGOTIPO (GIF):"
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   3570
         TabIndex        =   21
         Top             =   1785
         Width           =   1905
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SERIE"
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CheckBox CheqEstEmi 
      BackColor       =   &H00808080&
      Caption         =   "Asignar Usuario a Establecimiento o Punto de Emision"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   210
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   10
      Top             =   2730
      Width           =   7785
   End
   Begin InetCtlsObjects.Inet URLInet 
      Left            =   4305
      Top             =   5775
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox TextClaveOld 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2205
      Width           =   2010
   End
   Begin MSAdodcLib.Adodc AdoClave 
      Height          =   330
      Left            =   315
      Top             =   5775
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      Height          =   855
      Left            =   8190
      Picture         =   "IngClave.frx":9254
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   315
      Width           =   1065
   End
   Begin VB.TextBox TextNombreUsuario 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      MaxLength       =   40
      TabIndex        =   1
      Top             =   525
      Width           =   7785
   End
   Begin VB.TextBox TextUsuario 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2415
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2205
      Width           =   2010
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   8190
      Picture         =   "IngClave.frx":9696
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1470
      Width           =   1065
   End
   Begin VB.TextBox TextClave 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4620
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2205
      Width           =   2010
   End
   Begin MSAdodcLib.Adodc AdoMySQLClave 
      Height          =   330
      Left            =   2205
      Top             =   5775
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   2640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   4830
      TabIndex        =   8
      Top             =   1890
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Completo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2640
   End
   Begin VB.Label LabelUsuario 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   2625
      TabIndex        =   6
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "IngClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheqEstEmi_Click()
    If CheqEstEmi.value <> 0 Then FrmEstEmi.Visible = True Else FrmEstEmi.Visible = False
End Sub

Private Sub TextClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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
     sSQL = "SELECT * " _
          & "FROM Accesos " _
          & "WHERE Codigo = '" & TextClaveOld & "'"
     Select_Adodc AdoClave, sSQL
     With AdoClave.Recordset
      If .RecordCount > 0 Then
          MsgBox UCaseStrg("El Usuario ya registrado a " & vbCrLf & vbCrLf & .fields("Nombre_Completo"))
          TextClaveOld.SetFocus
      Else
          Command1.Enabled = True
          Command1.SetFocus
      End If
     End With
  Else
     If Tipo_RUC_CI.Tipo_Beneficiario <> "P" Then
        MsgBox "NUMERO DE CEDULA INCORRECTO," & vbCrLf & vbCrLf _
             & "VUELVA HA INGRESAR SIN GUIONES"
        Command1.Enabled = False
       'TextClaveOld.SetFocus
        TextEmail.SetFocus
     Else
        MsgBox "Esta es una NIC Extranjera"
        Command1.Enabled = True
     End If
  End If
End Sub

Private Sub Command1_Click()
Dim Serie As String

    sSQL = "SELECT Codigo " _
         & "FROM Accesos " _
         & "WHERE UCaseStrg(Usuario) = '" & UCaseStrg(TextUsuario.Text) & "' " _
         & "AND UCaseStrg(Clave) = '" & UCaseStrg(TextClave.Text) & "' "
    Select_Adodc AdoClave, sSQL
    If AdoClave.Recordset.RecordCount > 0 Then
       MsgBox "El Usuario: " & UCaseStrg(TextNombreUsuario.Text) & " ya existe. No se creara"
    Else
       Control_Procesos Normal, "Creacion de usuario: " & TextUsuario.Text
       Serie = TxtNumSerieUno & TxtNumSerieDos
       Codigo = TextClaveOld
       SetAdoAddNew "Accesos", True
       SetAdoFields "TODOS", True
       SetAdoFields "Clave", TextClave.Text
       SetAdoFields "Codigo", Codigo
       SetAdoFields "Usuario", TextUsuario.Text
       SetAdoFields "Nombre_Completo", TextNombreUsuario.Text
       SetAdoFields "EmailUsuario", TextEmail.Text
       If CheqEstEmi.value <> 0 Then SetAdoFields "Serie_FA", Serie
       SetAdoUpdate
       
      'SRI
       If CheqEstEmi.value <> 0 Then
          
          sSQL = "DELETE * " _
               & "FROM Catalogo_Lineas " _
               & "WHERE Codigo = 'S" & Serie & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' "
          Ejecutar_SQL_SP sSQL

          SetAdoAddNew "Catalogo_Lineas"
          SetAdoFields "Codigo", "S" & Serie
          SetAdoFields "Concepto", "CXC ELECTRONICOS " & Serie
          SetAdoFields "Fecha", FechaSistema
          SetAdoFields "Vencimiento", CLongFecha(CFechaLong(FechaSistema) + 365)
          SetAdoFields "Fact", "FA"
          SetAdoFields "Serie", Serie
          SetAdoFields "Secuencial", 1
          SetAdoFields "Autorizacion", RUC
          SetAdoFields "CxC", "1.1.03.01"
          SetAdoFields "CxC_Anterior", "1.1.03.01"
          SetAdoFields "Logo_Factura", "FactMult"
          SetAdoFields "Largo", 9.3
          SetAdoFields "Ancho", 10
          SetAdoFields "Nombre_Establecimiento", TextNombreUsuario.Text
          SetAdoFields "Direccion_Establecimiento", TxtDireccionEstab.Text
          SetAdoFields "Email_Establecimiento", TxtEmailEstEmi.Text
          SetAdoFields "Telefono_Estab", TxtTelefonoEstab.Text
          SetAdoFields "Logo_Tipo_Estab", TxtLogoTipoEstab.Text
          SetAdoFields "RUC_Establecimiento", Codigo & "001"
          SetAdoFields "TL", True
          SetAdoUpdate
       End If
       
      'Grabamos en Clientes Tambien
       sSQL = "SELECT Codigo " _
            & "FROM Clientes " _
            & "WHERE Codigo = '" & Codigo & "' "
       Select_Adodc AdoClave, sSQL
       If AdoClave.Recordset.RecordCount <= 0 Then
          SetAdoAddNew "Clientes", True
          SetAdoFields "Codigo", Codigo
          SetAdoFields "CI_RUC", Codigo
          SetAdoFields "TD", "C"
          SetAdoFields "Cliente", UCaseStrg(TextNombreUsuario.Text)
          SetAdoFields "Email", TextEmail.Text
          SetAdoFields "Grupo", NumEmpresa
          SetAdoUpdate
       End If
      
      'Grabamos en MySQL el mismo usuario
       sSQL = "SELECT CI_NIC, Nombre_Usuario, Usuario, Clave, TODOS, Supervisor, Tipo_Usuario, Email, Serie_FA, ID " _
            & "FROM acceso_usuarios " _
            & "WHERE CI_NIC = '" & Codigo & "' "
       Select_Adodc AdoMySQLClave, sSQL
       If AdoMySQLClave.Recordset.RecordCount <= 0 Then
          AdoMySQLClave.Recordset.AddNew
          AdoMySQLClave.Recordset.fields("CI_NIC") = Codigo
          AdoMySQLClave.Recordset.fields("Nombre_Usuario") = TextNombreUsuario.Text
          AdoMySQLClave.Recordset.fields("Usuario") = TextUsuario.Text
          AdoMySQLClave.Recordset.fields("Clave") = TextClave.Text
          AdoMySQLClave.Recordset.fields("TODOS") = True
          AdoMySQLClave.Recordset.fields("Supervisor") = True
          AdoMySQLClave.Recordset.fields("Tipo_Usuario") = "user"
          AdoMySQLClave.Recordset.fields("Email") = TextEmail.Text
          If CheqEstEmi.value <> 0 Then AdoMySQLClave.Recordset.fields("Serie_FA") = Serie
          AdoMySQLClave.Recordset.Update
       End If
    End If
    
    Titulo = "Pregunta de Creacion"
    Mensajes = "Desea Crear nuevo Usuario?"
    If BoxMensaje = vbYes Then
       TextEmail.Text = ""
       TextUsuario.Text = ""
       TextClave.Text = ""
       TextClaveOld.Text = ""
       TextNombreUsuario.Text = ""
       TextNombreUsuario.SetFocus
    Else
       Empresa = Ninguno
       Unload IngClave
       IngresarClave = True
       ListEmp.Show
    End If
End Sub

Private Sub Command2_Click()
  Unload IngClave
  IngresarClave = True
  ListEmp.Show
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm IngClave
  Redondear_Formulario IngClave, 35
  ConectarAdodc AdoClave
  ConectarAdodc_MySQL AdoMySQLClave
End Sub

Private Sub TextEmail_GotFocus()
   MarcarTexto TextEmail
End Sub

Private Sub TextEmail_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextEmail_LostFocus()
   TextoValido TextEmail
   TextEmail = LCase(TextEmail)
   TxtEmailEstEmi = TextEmail
End Sub

Private Sub TextNombreUsuario_GotFocus()
   MarcarTexto TextNombreUsuario
End Sub

Private Sub TextNombreUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TextNombreUsuario_LostFocus()
   TextNombreUsuario = ULCase(TextNombreUsuario)
End Sub

Private Sub TextUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub
