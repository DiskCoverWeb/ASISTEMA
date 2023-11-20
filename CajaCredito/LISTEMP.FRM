VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form ListEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPRESA A TRABAJAR"
   ClientHeight    =   4095
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Listemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2115
      Left            =   315
      Picture         =   "Listemp.frx":0442
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   14
      Top             =   210
      Width           =   5055
   End
   Begin VB.Data DataSeteos 
      Caption         =   "Seteos"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame FrameClave 
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   5475
      Begin VB.TextBox TextUsuario 
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
         TabIndex        =   3
         Top             =   315
         Width           =   1905
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
         Left            =   3990
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Clave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         TabIndex        =   2
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   1
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Crear &Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5670
      Picture         =   "Listemp.frx":31C9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2940
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Crear Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5670
      Picture         =   "Listemp.frx":360B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1995
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
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
      Height          =   960
      Left            =   5670
      Picture         =   "Listemp.frx":3A4D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1050
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5670
      Picture         =   "Listemp.frx":3CCF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1275
   End
   Begin VB.Data DataClave 
      Caption         =   "Clave"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3675
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame FrameEmpresa 
      Height          =   960
      Left            =   105
      TabIndex        =   5
      Top             =   2520
      Width           =   5475
      Begin MSDBCtls.DBCombo DBCEmpresa 
         Bindings        =   "Listemp.frx":4111
         DataSource      =   "DataEmpresa"
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   525
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   556
         _Version        =   327680
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TextDolar 
         Alignment       =   1  'Right Justify
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
         Left            =   3780
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "Listemp.frx":4127
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Elija la Empresa y presione [Aceptar]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cotización Dolar"
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
         Left            =   3780
         TabIndex        =   8
         Top             =   210
         Width           =   1590
      End
   End
   Begin VB.Data DataEmp 
      Caption         =   "Emp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3675
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Data DataEmpresa 
      Caption         =   "Empresa"
      Connect         =   "Access"
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
      Top             =   3675
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "ListEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Intentos As Integer
Dim Claves As String
Dim Tecla As Integer

Public Sub LlenarEmpresa()
Dim WrkJet As Workspace
Dim DataC As Database
Dim DataReg As Recordset
sSQL = "SELECT * FROM Empresas WHERE Empresa = '" & DBCEmpresa.Text & "' "
DataEmp.RecordSource = sSQL: DataEmp.Refresh
With DataEmp.Recordset
  If .RecordCount > 0 Then
      RatonReloj
      RutaEmpresa = ""
      RutaSubDirTemp = ""
      NumEmpresa = .Fields("Item")
      NumItemTemp = NumEmpresa
      OpcCoop = .Fields("Opc")
      CalcularIVA = .Fields("IVA")
      Empresa = .Fields("Empresa")
      Gerente = .Fields("Gerente")
      NombreCiudad = .Fields("Ciudad")
      NombrePais = .Fields("Pais")
      RUC = .Fields("RUC")
      FAX = .Fields("FAX")
      Moneda = .Fields("S_M")
      Telefono1 = .Fields("Telefono1")
      Telefono2 = .Fields("Telefono2")
      Direccion = .Fields("Direccion")
      Autorizacion = .Fields("Autorizacion")
      Fecha_Vence = .Fields("Vence")
      NombreRUC = "C.I.F."
      If .Fields("Logo_Tipo") = Ninguno Then
          LogoTipo = ""
      Else
          LogoTipo = UCase(RutaSistema & "\LOGOS\" & .Fields("Logo_Tipo") & ".WMF")
      End If
      AltoFactura = 11
      If CalcularIVA Then
         LogoFactura = UCase(RutaSistema & "\FORMATOS\FACTURAS.WMF")
      Else
         LogoFactura = UCase(RutaSistema & "\FORMATOS\NOTAVENT.WMF")
      End If
      RutaEmpresa = UCase(RutaSistema & "\EMPRESA\" & .Fields("SubDir"))
      RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA\" & .Fields("SubDir"))
      Dolar = CSng(TextDolar.Text)
      EmpresaActual = "[" & RutaEmpresa & "]."
      
      Set WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
      Set DataC = WrkJet.OpenDatabase(RutaEmpresa & "\SETEOS.MDB")
      sSQL = "UPDATE Cierre_Mes "
      sSQL = sSQL & "SET Fecha = #" & FechaSistema & "# "
      DataC.Execute sSQL
      sSQL = "SELECT * FROM Cierre_Mes "
      Set DataReg = DataC.OpenRecordset(sSQL, dbOpenDynaset, dbReadOnly)
      If DataReg.RecordCount > 0 Then
         FechaCierre = Format(DataReg.Fields("Fecha"), FormatoFechas)
      End If
      sSQL = "SELECT * FROM Cotizacion "
      SelectData DataSeteos, sSQL, False
      If DataSeteos.Recordset.RecordCount > 0 Then
         DataSeteos.Recordset.Edit
         DataSeteos.Recordset.Fields("Cotizacion") = Dolar
         DataSeteos.Recordset.Update
      End If
      SeteosCtas
      RatonNormal
  End If
End With
End Sub

Public Function BuscarClave(Usuario As String, Clave As String) As Boolean
Dim Respuesta As Boolean
  NombreUsuario = "": Respuesta = True
  sSQL = "SELECT * FROM Accesos WHERE Ucase(Usuario) = '" & UCase(Usuario) & "' "
  DataClave.RecordSource = sSQL: DataClave.Refresh
  With DataClave.Recordset
   If .RecordCount > 0 Then
      If UCase(.Fields("Clave")) = UCase(Clave) Then
       'Or (.Fields(TipoModulo) = True)
         If .Fields("TODOS") = True Then
            NombreUsuario = .Fields("Nombre_Completo")
            CodigoUsuario = .Fields("Codigo")
         ElseIf .Fields(TipoModulo) = False Then
            Respuesta = False
            Cadena = "Warning: Usted no esta autorizado " & Chr(13)
            Cadena = Cadena & Space(10) & "a ingresar a este modulo."
            MsgBox Cadena
         End If
      Else
         Respuesta = False
         MsgBox "Error: Clave incorrecta."
      End If
   Else
     Respuesta = False
     Cadena = "Sr(a). " & Usuario & ": " & Chr(13)
     Cadena = Cadena & Space(10) & "Usted no esta autorizado a ingresar al sistema." & Chr(13)
     If Intentos <= 3 Then Cadena = Cadena & Space(10) & "Vuelva a ingresar su clave."
     MsgBox Cadena
   End If
  End With
  sSQL = "SELECT * FROM " & App.EXEName & " "
  sSQL = sSQL & "WHERE Codigo = '" & CodigoUsuario & "' "
  DataClave.RecordSource = sSQL: DataClave.Refresh
  With DataClave.Recordset
   If .RecordCount > 0 Then
       CNivel_1 = .Fields("Nivel_1")
       CNivel_2 = .Fields("Nivel_2")
       CNivel_3 = .Fields("Nivel_3")
       CNivel_4 = .Fields("Nivel_4")
       CNivel_5 = .Fields("Nivel_5")
       CNivel_6 = .Fields("Nivel_6")
       Supervisor = .Fields("Supervisor")
   End If
  End With
  BuscarClave = Respuesta
End Function

Private Sub Command1_Click()
   End
End Sub

Private Sub Command2_Click()
  If IngresarClave = False Then
     FrameClave.Visible = False
     FrameEmpresa.Visible = True
     LlenarEmpresa
     If RutaEmpresa <> "" Then
        ChDir RutaEmpresa
        Unload ListEmp
        IniciarPrograma = True
     End If
  End If
End Sub

Private Sub Command3_Click()
 If ClaveSupervisor Then
    Unload ListEmp
    IngClave.Show
 End If
End Sub

Private Sub Command4_Click()
  If ClaveSupervisor Then
     Unload ListEmp
     CrearEmp.Show
  End If
End Sub

Private Sub DBCEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub Form_Activate()
   CNivel_1 = True
   CNivel_2 = True
   CNivel_3 = True
   CNivel_4 = True
   CNivel_5 = True
   CNivel_6 = True
   Supervisor = True
   Intentos = 0: ClaveGeneral = ""
   ListEmp.Caption = "UNIDAD DE RED: [" & RutaSistema & "]."
   sSQL = "SELECT * FROM Accesos WHERE Ucase(Usuario) = 'ADMINISTRA' "
   SelectData DataClave, sSQL, False
   With DataClave.Recordset
    If .RecordCount > 0 Then ClaveGeneral = .Fields("Clave")
   End With
   sSQL = "SELECT * FROM Cotizacion "
   DataSeteos.RecordSource = sSQL: DataSeteos.Refresh
   If DataSeteos.Recordset.RecordCount > 0 Then
      TextDolar.Text = Format(DataSeteos.Recordset.Fields("Cotizacion"), "#,###.00")
   Else
      TextDolar.Text = "0"
   End If
   sSQL = "SELECT * FROM Empresas ORDER BY Empresa,Item "
   'sSQL = "SELECT * FROM Empresas ORDER BY Empresa "
   SelectDBCombo DBCEmpresa, DataEmpresa, sSQL, "Empresa", False
   If DataEmpresa.Recordset.RecordCount <= 0 Then
      Unload ListEmp
      CrearEmp.Show
   End If
   If IngresarClave Then
      FrameClave.Visible = True
      FrameEmpresa.Visible = False
      Command2.Enabled = False
      Command3.Enabled = False
      Command4.Enabled = False
      TextUsuario.SetFocus
   Else
      IniciarPrograma = True
      FrameClave.Visible = False
      FrameEmpresa.Visible = True
      DBCEmpresa.SetFocus
   End If
End Sub

Private Sub Form_Load()
   CentrarForm ListEmp
   DataEmp.DatabaseName = RutaSistema & "\EMPRESAS.MDB"
   DataEmpresa.DatabaseName = RutaSistema & "\EMPRESAS.MDB"
   DataClave.DatabaseName = RutaSistema & "\EMPRESAS.MDB"
   DataSeteos.DatabaseName = RutaSistema & "\SETEOS.MDB"
  'Obtenemos la fecha del Sistema
   FechaSistema = Format(Day(Date), "00") & "/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
   Dia = Format(Day(Date), "00")
   Mes = Format(Month(Date), "00")
   Anio = Format(Year(Date), "0000")
   H_INCH = 1440 / Screen.TwipsPerPixelX
   V_INCH = 1440 / Screen.TwipsPerPixelY
End Sub

Private Sub TextDolar_GotFocus()
   MarcarTexto TextDolar
   IniciarPrograma = True
End Sub

Private Sub TextDolar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Key_Enter Then
     LlenarEmpresa
     If RutaEmpresa <> "" Then
        ChDir RutaEmpresa
        Unload ListEmp
     End If
  End If
End Sub

Private Sub TextDolar_LostFocus()
   TextDolar.Text = Format(TextDolar.Text, "#,##0.00")
End Sub

Private Sub TextClave_Change()
  If Len(TextClave.Text) >= TextClave.MaxLength Then SendKeys "{TAB}"
End Sub

Private Sub TextClave_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextClave_LostFocus()
Claves = TextClave.Text
If Claves <> "" And TextUsuario.Text <> "" Then
   Intentos = Intentos + 1
   If (BuscarClave(TextUsuario.Text, Claves)) And (Intentos < 3) Then
      FrameClave.Visible = False
      FrameEmpresa.Visible = True
      Command2.Enabled = True
      Command3.Enabled = True
      Command4.Enabled = True
      IngresarClave = False
      DBCEmpresa.SetFocus
   ElseIf Intentos >= 3 Then
      Cadena = "Sr(a). " & UCase(TextUsuario.Text) & ": " & Chr(13)
      Cadena = Cadena & Space(10) & "Usted no está autorizado" & Chr(13)
      Cadena = Cadena & Space(10) & "a ingresar al sistema." & Chr(13) & Chr(13)
      Cadena = Cadena & Space(10) & "Vuelva a ejecutar el programa."
      MsgBox Cadena
      End
   Else
      Claves = "": TextClave.Text = "": TextClave.SetFocus
   End If
End If
End Sub

Private Sub TextUsuario_GotFocus()
  TextUsuario.Text = ""
End Sub

Private Sub TextUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextUsuario_LostFocus()
  Claves = "": TextClave.Text = ""
End Sub

