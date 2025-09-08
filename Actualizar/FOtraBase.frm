VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FOtraBase 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECCIONE LA ENTIDAD A CONECTAR"
   ClientHeight    =   1980
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7500
   DrawMode        =   14  'Copy Pen
   Icon            =   "FOtraBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   7500
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6720
      Picture         =   "FOtraBase.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6720
      Picture         =   "FOtraBase.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   645
   End
   Begin VB.TextBox TxtReferencia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1050
      Width           =   6525
   End
   Begin MSDataListLib.DataList DLEntidad 
      Bindings        =   "FOtraBase.frx":1A5E
      DataSource      =   "AdoEntidad"
      Height          =   840
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   1482
      _Version        =   393216
      BackColor       =   16761087
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoEntidad 
      Height          =   330
      Left            =   315
      Top             =   3570
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
      Caption         =   "Entidad"
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
Attribute VB_Name = "FOtraBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim AdoStrCnnTemp As String
  AdoStrCnnTemp = AdoStrCnn
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
         'Determinar que tipo de bases que utilizamos
          NumEmpresa = "999"
          Si_No = False
          Evaluar = False
          Modo_Educativo = False
          SQL_Server = False
          
          strIPServidor = .Fields("IP_VPN_RUTA")
          strNombreBaseDatos = .Fields("Base_Datos")
          strWebServices = .Fields("WebServices")
          strPassword = .Fields("Clave_DB")
          strUsuario = .Fields("Usuario_DB")
          strPuerto = .Fields("Puerto")
          
          Select Case .Fields("Tipo_Base")
            Case "SQL SERVER"
                 If .Fields("Puerto") <> 1433 Then
                     AdoStrCnn = "Data Source=tcp:" & strIPServidor & "," & CStr(strPuerto) & ";" & vbCrLf
                 Else
                     AdoStrCnn = "Data Source=" & strIPServidor & ";" & vbCrLf
                 End If
                 AdoStrCnn = AdoStrCnn _
                           & "Initial Catalog=" & strNombreBaseDatos & ";" & vbCrLf _
                           & "Provider=SQLOLEDB.1;" & vbCrLf _
                           & "UID=" & strUsuario & ";" & vbCrLf _
                           & "PWD=" & strPassword & ";"
                 If Len(strWebServices) = 3 Then AdoStrCnn = AdoStrCnn & vbCrLf & "WebServices=" & strWebServices & "; "
                 SQL_Server = True
            Case "MY SQL"
                 AdoStrCnn = "SERVER=" & strIPServidor & ";" & vbCrLf _
                           & "DATABASE=" & strNombreBaseDatos & ";" & vbCrLf _
                           & "DRIVER={MySQL ODBC 3.51 Driver};" & vbCrLf _
                           & "UID=" & strUsuario & ";" & vbCrLf _
                           & "PWD=" & strPassword & ";" & vbCrLf _
                           & "OPTION=3;"
            Case "ACCESS"
                 AdoStrCnn = "Data Source=" & strIPServidor & "\" & strNombreBaseDatos & ".MDB;" & vbCrLf _
                           & "Provider=Microsoft.Jet.OLEDB.4.0;" & vbCrLf _
                           & "Persist Security Info=False;"
          End Select

          If Not Ping_IP(strIPServidor) Then
             MsgBox "LA CONEXION NO ESTA ESTABLECIDA" & vbCrLf _
                  & "POR FAVOR LLAME AL BENEFICIARIO" & vbCrLf _
                  & "PARA QUE CONECTE LA VPN"
             AdoStrCnn = AdoStrCnnTemp
          Else
             Dolar = 0
             ConSucursal = False
             FUpDateExec.Caption = Modulo & ": " & strIPServidor & " - " & strNombreBaseDatos
             RatonNormal
          End If
''          MsgBox strIPServidor & vbCrLf _
''               & strNombreBaseDatos & vbCrLf _
''               & strWebServices & vbCrLf _
''               & strPassword & vbCrLf _
''               & strUsuario & vbCrLf _
''               & strPuerto & vbCrLf
       Else
          MsgBox "No ha seleccionado ninguna Entidad"
       End If
   Else
      MsgBox "No hay Entidad asignada"
   End If
  End With
  RatonNormal
  Unload FOtraBase
End Sub

Private Sub Command2_Click()
   Unload FOtraBase
End Sub

Private Sub Form_Activate()
    sSQL = "SELECT * " _
         & "FROM Empresas_Externas " _
         & "WHERE Entidad_Comercial <> '.' " _
         & "ORDER BY Entidad_Comercial "
    SelectDB_List DLEntidad, AdoEntidad, sSQL, "Entidad_Comercial"
    If AdoEntidad.Recordset.RecordCount > 0 Then
       DLEntidad.SetFocus
    Else
       MsgBox "LLAME A SU PROVEEDOR PARA QUE CONFIGURE" & vbCrLf _
            & "ESTA OPCION Y PODER DISFRUTAR LA NUEVA" & vbCrLf _
            & "FORMA DE CONECTARCE CON OTRAS ENTIDADES" & vbCrLf _
            & "EMAIL: diskcover@msn.com" & vbCrLf _
            & "diskcoversystem@msn.com" & vbCrLf _
            & "TELEFONO BPX: 593-02-3210-051" & vbCrLf
       Unload FOtraBase
    End If
End Sub

Private Sub Form_Load()
   'CentrarForm FOtraBase
   ConectarAdodc AdoEntidad
End Sub

Private Sub DLEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
Dim buscarEmpresa As String
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyB Then
     buscarEmpresa = InputBox("Patron de Busqueda por Entidad", "BUSCAR POR ENTIDAD", "")
     sSQL = "SELECT * " _
          & "FROM Empresas_Externas " _
          & "WHERE Entidad_Comercial LIKE '%" & buscarEmpresa & "%' " _
          & "ORDER BY Entidad_Comercial "
     SelectDB_List DLEntidad, AdoEntidad, sSQL, "Entidad_Comercial"
     If AdoEntidad.Recordset.RecordCount <= 0 Then
        MsgBox "LLAME A SU PROVEEDOR PARA QUE CONFIGURE ESTA OPCION Y PODER" & vbCrLf _
             & "DISFRUTAR LA NUEVA FORMA DE CONECTARCE CON OTRAS ENTIDADES." & vbCrLf _
             & vbCrLf _
             & "EMAIL: diskcover@msn.com o diskcoversystem@msn.com" & vbCrLf & vbCrLf _
             & "TELEFONO BPX: 593-02-6052430" & vbCrLf
     End If
     DLEntidad.SetFocus
  End If
  If KeyCode = vbKeyEscape Then Unload FOtraBase
End Sub

Private Sub DLEntidad_KeyUp(KeyCode As Integer, Shift As Integer)
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
          TxtReferencia = "IP VPN    : " & .Fields("IP_VPN_RUTA") & vbCrLf _
                        & "BASE DATOS: " & .Fields("Base_Datos") & vbCrLf _
                        & "USUARIO   : " & .Fields("Usuario_DB") & vbTab & "CLAVE DB: " & .Fields("Clave_DB")
       End If
   End If
  End With
End Sub

