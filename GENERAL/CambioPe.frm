VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form CambioPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAMBIO DE PERIODO"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   Icon            =   "CambioPe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstPeriodo 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "CambioPe.frx":030A
      Left            =   105
      List            =   "CambioPe.frx":030C
      TabIndex        =   1
      Top             =   420
      Width           =   4320
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
      Height          =   960
      Left            =   4725
      Picture         =   "CambioPe.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1155
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cambiar Periodo"
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
      Left            =   4725
      Picture         =   "CambioPe.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoPeriodo 
      Height          =   330
      Left            =   105
      Top             =   1995
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Periodo"
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
   Begin MSAdodcLib.Adodc AdoEmp 
      Height          =   330
      Left            =   210
      Top             =   630
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
      Caption         =   "Emp"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccione el &Periodo"
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
      Width           =   4320
   End
End
Attribute VB_Name = "CambioPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim IdP As Integer
Dim IdPS As Integer

  Periodo_Contable = LstPeriodo.Text
  If Periodo_Contable = "Periodo Actual" Then Periodo_Contable = Ninguno
  If Periodo_Contable <> Ninguno Then
     Periodo_Superior = Periodo_Contable
     IdPS = 0
     For IdP = 0 To LstPeriodo.ListCount - 1
         If Periodo_Contable = LstPeriodo.List(IdP) Then IdPS = IdP
     Next IdP
     If IdPS > 0 Then Periodo_Superior = LstPeriodo.List(IdPS - 1)
  Else
     Periodo_Superior = Ninguno
  End If
  If Periodo_Superior = "Periodo Actual" Then Periodo_Superior = Ninguno
  Anio_Lectivo = Ninguno
  Director = Ninguno: Secretario1 = Ninguno
  Rector = Ninguno:   Secretario2 = Ninguno
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_Adodc AdoEmp, sSQL
  With AdoEmp.Recordset
   If .RecordCount > 0 Then
       Anio_Lectivo = .Fields("Anio_Lectivo")
       Director = .Fields("Director")
       Secretario1 = .Fields("Secretario1")
       Rector = .Fields("Rector")
       Secretario2 = .Fields("Secretario2")
   End If
  End With
  Crear_Cierre_Mes
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  RatonReloj
  LstPeriodo.Clear
  LstPeriodo.AddItem "Periodo Actual"
  Periodo_Superior = Periodo_Contable
  sSQL = "SELECT Periodo "
  If UCaseStrg(Modulo) = "EDUCATIVO" Then
     sSQL = sSQL & "FROM Catalogo_Periodo_Lectivo "
  ElseIf UCaseStrg(Modulo) = "FACTURACION" Then
     sSQL = sSQL & "FROM Facturas "
  Else
     sSQL = sSQL & "FROM Catalogo_Cuentas "
  End If
  sSQL = sSQL & "WHERE Periodo <> '.' "
  If ConSucursal = False Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "GROUP BY Periodo " _
       & "ORDER BY Periodo "
  Select_Adodc AdoPeriodo, sSQL
  With AdoPeriodo.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstPeriodo.AddItem CStr(.Fields("Periodo"))
         .MoveNext
       Loop
   Else
      Cadena = "No hay información que procesar, si esta seguro" & vbCrLf & vbCrLf _
             & "de subir datos ingrese el periodo a subir" & vbCrLf & vbCrLf & vbCrLf _
             & "INGRESE EL PERIODO"
      Periodo_Contable = InputBox(Cadena, "CAMBIOS DE PERIODOS", ".")
      LstPeriodo.AddItem Periodo_Contable
   End If
  End With
  RatonNormal
  LstPeriodo.Text = LstPeriodo.List(0)
  LstPeriodo.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm CambioPeriodo
  ConectarAdodc AdoEmp
  ConectarAdodc AdoPeriodo
End Sub

Private Sub LstPeriodo_DblClick()
  SiguienteControl
End Sub

Private Sub LstPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload Me
  If CtrlDown And KeyCode = vbKeyP Then
     Periodo_Contable = InputBox(Cadena, "CAMBIOS DE PERIODOS", ".")
     LstPeriodo.AddItem Periodo_Contable
  End If
End Sub
