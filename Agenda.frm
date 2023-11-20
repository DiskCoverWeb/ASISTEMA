VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FAgenda 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AGENDA ELECTRONICA"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7785
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
   Icon            =   "Agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DCClientes 
      Bindings        =   "Agenda.frx":164A
      DataSource      =   "AdoClientes"
      Height          =   1740
      Left            =   105
      TabIndex        =   2
      ToolTipText     =   "<CTRL+B> Buscar datos"
      Top             =   315
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3069
      _Version        =   393216
      Style           =   1
      BackColor       =   16761024
      ForeColor       =   8388608
      Text            =   ""
   End
   Begin VB.Frame FrmBusqueda 
      BackColor       =   &H00C0FFC0&
      Caption         =   "|PATRON DE BUSQUEDA|"
      Height          =   645
      Left            =   105
      TabIndex        =   35
      Top             =   6825
      Visible         =   0   'False
      Width           =   5370
      Begin VB.TextBox TxtBusqueda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   36
         Top             =   210
         Width           =   5160
      End
   End
   Begin VB.Frame FrmDatos 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DATOS DEL BENEFICIARIOS"
      ForeColor       =   &H00FFFFFF&
      Height          =   4635
      Left            =   105
      TabIndex        =   4
      Top             =   2100
      Width           =   7575
      Begin VB.TextBox TxtProvincia 
         Height          =   330
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   38
         Top             =   3570
         Width           =   4005
      End
      Begin VB.TextBox TxtNacionalidad 
         Height          =   330
         Left            =   105
         MaxLength       =   50
         TabIndex        =   39
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox TxtCiudad 
         Height          =   330
         Left            =   105
         MaxLength       =   50
         TabIndex        =   37
         Top             =   4200
         Width           =   3375
      End
      Begin VB.TextBox TxtDirS 
         Height          =   330
         Left            =   105
         MaxLength       =   60
         TabIndex        =   23
         Top             =   1695
         Width           =   7365
      End
      Begin VB.TextBox TxtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   6300
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Agenda.frx":1664
         Top             =   2310
         Width           =   1170
      End
      Begin VB.TextBox TxtContacto 
         Height          =   330
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2310
         Width           =   5055
      End
      Begin VB.TextBox TxtCelular 
         Height          =   330
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "9999999999"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtTelefonoT 
         Height          =   330
         Left            =   3885
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "9999999999"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtFAX 
         Height          =   330
         Left            =   2730
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "9999999999"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtTelefonoS 
         Height          =   330
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "9999999999"
         Top             =   1050
         Width           =   1170
      End
      Begin VB.TextBox TxtApellidosS 
         Height          =   330
         Left            =   105
         MaxLength       =   60
         TabIndex        =   21
         Top             =   420
         Width           =   7365
      End
      Begin VB.TextBox TxtNumero 
         Height          =   330
         Left            =   105
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2310
         Width           =   1170
      End
      Begin VB.TextBox TxtRazonSocial 
         Height          =   330
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   26
         Top             =   4200
         Width           =   4005
      End
      Begin VB.TextBox TxtEmail 
         Height          =   330
         Left            =   105
         MaxLength       =   120
         TabIndex        =   24
         Top             =   2940
         Width           =   7365
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   285
         Left            =   6090
         MaxLength       =   60
         TabIndex        =   22
         Top             =   210
         Width           =   1380
      End
      Begin VB.TextBox TxtCI_RUC 
         Height          =   330
         Left            =   105
         MaxLength       =   13
         TabIndex        =   12
         Text            =   "9999999999999"
         ToolTipText     =   "<Alt+F2> Codigo Automático"
         Top             =   1050
         Width           =   1485
      End
      Begin VB.TextBox TxtGrupo 
         Height          =   330
         Left            =   6195
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* PROVINCIA"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3465
         TabIndex        =   31
         Top             =   3360
         Width           =   4005
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* CIUDAD"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   29
         Top             =   4005
         Width           =   3375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   6300
         TabIndex        =   13
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CONTACTO"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   1260
         TabIndex        =   14
         Top             =   2100
         Width           =   5055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CELULAR"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   5040
         TabIndex        =   16
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3885
         TabIndex        =   15
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FAX"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   2730
         TabIndex        =   18
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TELEFONO"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   1575
         TabIndex        =   19
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* NUMERO"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   28
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* APELLIDOS Y NOMBRES"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   34
         Top             =   210
         Width           =   6000
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " REPRESENTANTE"
         Height          =   225
         Left            =   3465
         TabIndex        =   33
         Top             =   3990
         Width           =   4005
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* NACIONALIDAD"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   32
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* DIRECCION"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   30
         Top             =   1470
         Width           =   7365
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMAIL (CORREO ELECTRONICO)"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   27
         Top             =   2730
         Width           =   7365
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " C.I./R.U.C."
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   20
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " GRUPO #"
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   6195
         TabIndex        =   17
         Top             =   840
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5670
      Top             =   6825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5355
      Top             =   6825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "&Ocultar"
      Height          =   645
      Left            =   6300
      TabIndex        =   1
      Top             =   6825
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3570
      Picture         =   "Agenda.frx":1668
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   630
      Width           =   1020
   End
   Begin MSAdodcLib.Adodc AdoAgenda 
      Height          =   330
      Left            =   210
      Top             =   525
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Agenda"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   210
      Top             =   840
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoProvincia 
      Height          =   330
      Left            =   210
      Top             =   1155
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Provincia"
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BUSQUE EL BENEFICIARIO"
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   7575
   End
End
Attribute VB_Name = "FAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Primera_Vez As Boolean
Dim PatronBusqueda As String
Dim cDestino As String
Const ContraseñaDiskCover = "diskcover1210"

Public Sub Listar_Datos_Cliente(CodigoCli As String)
    TxtCI_RUC = ""
    TxtCodigo = ""
    TxtApellidosS = ""
    TxtTelefonoS = ""
    TxtTelefonoT = ""
    TxtCelular = ""
    TxtFAX = ""
    TxtGrupo = ""
    TxtDirS = ""
    TxtNumero = ""
    TxtContacto = ""
    TxtDescuento = ""
    TxtRazonSocial = ""
    TxtEmail = ""
    sSQL = "SELECT FA, Cliente, TD, CI_RUC, Ciudad, Telefono, TelefonoT, FAX, Celular, Email, Email2, Direccion, " _
         & "DirNumero, Ciudad, Prov, Pais, Codigo, Sexo, Descuento, Contacto, Grupo, Representante " _
         & "FROM Clientes " _
         & "WHERE Codigo = '" & CodigoCli & "' " _
         & "ORDER BY Cliente "
    SelectAdodc AdoAgenda, sSQL
    With AdoAgenda.Recordset
     If .RecordCount > 0 Then
         TxtCI_RUC = .Fields("CI_RUC")
         TxtCodigo = .Fields("Codigo")
         TxtApellidosS = .Fields("Cliente")
         TxtTelefonoS = .Fields("Telefono")
         TxtTelefonoT = .Fields("TelefonoT")
         TxtCelular = .Fields("Celular")
         TxtFAX = .Fields("FAX")
        
         TxtGrupo = .Fields("Grupo")
         TxtDirS = .Fields("Direccion")
         TxtNumero = .Fields("DirNumero")
         TxtContacto = .Fields("Contacto")
         TxtDescuento = .Fields("Descuento")
         TxtRazonSocial = .Fields("Representante")
         TxtNacionalidad = "ECUADOR"
         TxtProvincia = .Fields("Prov")
         TxtCiudad = .Fields("Ciudad")
         If Len(.Fields("Email")) > 1 Then TxtEmail = .Fields("Email")
         If Len(.Fields("Email2")) > 1 Then
            If Len(TxtEmail) > 1 Then
               TxtEmail = TxtEmail & "; " & .Fields("Email2")
            Else
               TxtEmail = TxtEmail & .Fields("Email2")
            End If
         End If
         sSQL = "SELECT * " _
              & "FROM Tabla_Naciones " _
              & "WHERE CProvincia = '" & TxtProvincia & "' " _
              & "AND TR = 'P' "
         SelectAdodc AdoProvincia, sSQL
         If AdoProvincia.Recordset.RecordCount > 0 Then
            TxtProvincia = AdoProvincia.Recordset.Fields("Descripcion_Rubro")
         End If
     Else
         TxtBusqueda = "No existen datos"
     End If
    End With
End Sub

Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hWnd = Picture1.hWnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "Agenda Telefonica DiskCover System" & Chr$(0)
    Erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
 
Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hWnd = Picture1.hWnd
    Tic.uID = 1&
    Erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub Command1_Click()
   CreateIcon
   FAgenda.Visible = False
End Sub

Private Sub DCClientes_Change()
   PatronBusqueda = DCClientes
End Sub

Private Sub DCClientes_DblClick(Area As Integer)
    SiguienteControl
End Sub

Private Sub DCClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyQ Then
      'DeleteIcon
       End
    End If
    If CtrlDown And KeyCode = vbKeyB Then
       sSQL = "SELECT Cliente,Codigo " _
            & "FROM Clientes "
       If Len(DCClientes) > 1 Then
          sSQL = sSQL & "WHERE Cliente LIKE '%" & DCClientes & "%' "
       Else
          sSQL = sSQL & "WHERE Cliente <> '.' "
       End If
       sSQL = sSQL & "ORDER BY Cliente "
       SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
    Else
    
    End If
    PresionoEnter KeyCode
End Sub

Private Sub DCClientes_LostFocus()
Dim Codigo As String
    Codigo = Ninguno
    With AdoClientes.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & DCClientes & "' ")
         If Not .EOF Then Codigo = .Fields("Codigo")
     End If
    End With
    Listar_Datos_Cliente Codigo
End Sub

Private Sub Form_Activate()
   'Presentamos datos de la agenda
    sSQL = "SELECT Cliente,Codigo " _
         & "FROM Clientes " _
         & "WHERE Cliente <> '.' " _
         & "ORDER BY Cliente "
    SelectDBCombo DCClientes, AdoClientes, sSQL, "Cliente"
    RatonNormal
    DCClientes.SetFocus
End Sub

Private Sub Form_Load()
Dim Idx As Integer
Dim ContadorTime As Long
Dim CrearBaseDatos As Boolean
Dim LineaFile As Byte
Dim RutaFile As String
Dim Txt_SMTP_Mails As String

    MDI_Y_Max = Screen.Height
    MDI_X_Max = Screen.Width
    
    CentrarForm FAgenda
    
    Email_Respaldo = Ninguno
    Modulo = "AGENDA"

    Primera_Vez = True
    FechaSistema = Format(date, FormatoFechas)
    TiempoSistema = Time
    
   'Intervalo de espera antes de empezar a sacar los respaldos
    Timer1.Enabled = True
    Timer1.Interval = 1000  '1/2 segundo
    RatonReloj
    Unidad = Left(CurDir$, 2)
    RutaDestino = Unidad & "\SISTEMA"
    RutaSistema = Unidad & "\SISTEMA"
    RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
    RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
    RutaSysBases = Unidad & "\SYSBASES"
    RutaUpdate = RutaDestino
    ChDir RutaDestino
    
   'Determinar que tipo de bases utilizamos
    Evaluar = False
    SQL_Server = True
    Conectar_Base_Datos
   'Determinamos si vamos a utilizar la opcion de envio de email
    File_Emails = False
    Cadena = Dir(RutaSistema & "\", vbNormal) 'Recupera la primera entrada.
    Do While Cadena <> ""
       If Cadena <> "." And Cadena <> ".." Then
          If (GetAttr(RutaSistema & "\" & Cadena) And vbNormal) = vbNormal Then
             If UCase(Cadena) = "SMTP_MAILS.TXT" Then File_Emails = True
          End If
       End If
       Cadena = Dir
    Loop
   'Sacamos los datos del remitente de correos electronicos
    With TMail
        .ListaMail = 0
         Lista_De_Correos(0).Correo_Electronico = "diskcover.system@gmail.com"
         Lista_De_Correos(0).Contraseña = ContraseñaDiskCover
    End With
   'MsgBox Weekday(FechaSistema)
    NumEmpresa = "000"
    CodigoUsuario = "ACCESO01"
    NombreUsuario = "Supervisor General"
    Empresa = "MODULO DE ACTUALIZACION DE BASES Y DATOS"
    Periodo_Contable = "."
    
    RatonReloj
    ConectarAdodc AdoAgenda
    ConectarAdodc AdoClientes
    ConectarAdodc AdoProvincia
   
   'Averiguamos el mail de respaldo
    sSQL = "SELECT Item,Empresa,Email_Respaldos,RUC " _
         & "FROM Empresas " _
         & "WHERE LEN(Email_Respaldos) > 1 " _
         & "ORDER BY Item,Empresa,Email_Respaldos,RUC "
    SelectAdodc AdoAgenda, sSQL
    If AdoAgenda.Recordset.RecordCount > 0 Then Email_Respaldo = AdoAgenda.Recordset.Fields("Email_Respaldos")
    FAgenda.Caption = "AGENDA DE CONTACTOS"
    CreateIcon
    FAgenda.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  x = x / Screen.TwipsPerPixelX
  Select Case x
    Case WM_LBUTTONDOWN
         Caption = "Left Click"
    Case WM_RBUTTONDOWN
         Caption = "Right Click"
    Case WM_MOUSEMOVE
         Caption = "Move"
    Case WM_LBUTTONDBLCLK
         Caption = "Double Click"
         FAgenda.Visible = True
  End Select
End Sub

Private Sub Timer1_Timer()
  Minutos = Minute(Time - TiempoSistema)
  Segundos = Second(Time - TiempoSistema)
  MiTiempo = CSng(Format(Minutos, "00") & "." & Format(Segundos, "00"))
  FrmDatos.Caption = " FECHA: " & UCase$(Format(date, "dd/MMMM/yyyy")) & " " & String(35, "-") & " HORA ACTUAL: " & Format(Time, "HH:MM:SS") & " "
End Sub

