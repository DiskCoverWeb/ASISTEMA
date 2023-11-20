VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCopyEmpresa 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COPIAR CATALOGO DE OTRAS EMPRESAS"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmEntidad 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SELECCIONE LA ENTIDAD ORIGEN A COPIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4845
      Left            =   4515
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   6525
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
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
         Height          =   855
         Left            =   5355
         Picture         =   "FCopyEmp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2835
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
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
         Left            =   5355
         Picture         =   "FCopyEmp.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3780
         Width           =   1065
      End
      Begin VB.TextBox TxtReferencia 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1800
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2835
         Width           =   5160
      End
      Begin MSDataListLib.DataList DLEntidad 
         Bindings        =   "FCopyEmp.frx":1194
         DataSource      =   "AdoEntidad"
         Height          =   2400
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4233
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   65280
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
   End
   Begin VB.ListBox LstTablas 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   315
      Width           =   3900
   End
   Begin VB.CheckBox CheqBorrarEmp 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Borrar Base Datos destino"
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
      Left            =   4095
      TabIndex        =   2
      Top             =   4935
      Width           =   2745
   End
   Begin MSDataListLib.DataList DLEmpresa 
      Bindings        =   "FCopyEmp.frx":11AD
      DataSource      =   "AdoEmp"
      Height          =   4545
      Left            =   4095
      TabIndex        =   1
      Top             =   315
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   8017
      _Version        =   393216
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
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   10710
      Picture         =   "FCopyEmp.frx":11C2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   960
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
      Left            =   10710
      Picture         =   "FCopyEmp.frx":1A8C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1050
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   315
      Top             =   840
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Cta"
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
      Left            =   315
      Top             =   1155
      Width           =   2430
      _ExtentX        =   4286
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
   Begin MSAdodcLib.Adodc AdoCopyOrigen 
      Height          =   330
      Left            =   315
      Top             =   1470
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "CopyOrigen"
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
   Begin MSAdodcLib.Adodc AdoEntidad 
      Height          =   330
      Left            =   315
      Top             =   1785
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TABLAS A COPIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   3900
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA EMPRESA A COPIAR EL CATALOGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4095
      TabIndex        =   0
      Top             =   105
      Width           =   6525
   End
End
Attribute VB_Name = "FCopyEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim RegInsertarDB As ADODB.Recordset
Dim Existe_Fecha  As Boolean
Dim Existe_Codigo As Boolean
Dim NombreTabla   As String
 
  Cadena = DLEmpresa.Text
  NumItem = Ninguno
  If Cadena = "" Then Cadena = Ninguno
  With AdoEmp.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Empresa LIKE '" & Cadena & "' ")
       If Not .EOF Then NumItem = .Fields("Item")
   End If
  End With
  
'''  MsgBox AdoCopyOrigen.ConnectionString & vbCrLf _
'''       & NumItem & vbCrLf _
'''       & String(80, "-") & vbCrLf _
'''       & AdoAux.ConnectionString & vbCrLf _
'''       & NumEmpresa
  If NumItem <> Ninguno Then
     Mensajes = "Quiere Migrar la informacion de la Empresa" & vbCrLf _
              & DLEmpresa & " (" & NumItem & ")." & vbCrLf _
              & "a la Empresa " & Empresa & " (" & NumEmpresa & ")." & vbCrLf _
              & "Este proceso reemplazará la informacion en la empresa actual."
     Titulo = "Pregunta de Copia"
     If BoxMensaje = vbYes Then
        For i = 0 To LstTablas.ListCount - 1
            If LstTablas.Selected(i) Then
               Existe_Fecha = False
               Existe_Codigo = False
               NombreTabla = LstTablas.Selected(i)
              'Leemos datos origen
               sSQL = "SELECT TOP 1 * " _
                    & "FROM " & NombreTabla & " " _
                    & "WHERE Item = '" & NumItem & "' "
               Select_Adodc AdoCopyOrigen, sSQL
               For J = 0 To AdoCopyOrigen.Recordset.Fields.Count - 1
                   If AdoCopyOrigen.Recordset.Fields(J).Name = "Fecha" Then Existe_Fecha = True
                   If AdoCopyOrigen.Recordset.Fields(J).Name = "Codigo" Then Existe_Codigo = True
               Next J
               
               sSQL = "SELECT * " _
                    & "FROM " & NombreTabla & " " _
                    & "WHERE Item = '" & NumItem & "' "
               If Existe_Fecha And Existe_Codigo Then sSQL = sSQL & "ORDER BY Fecha, Codigo "
               If Existe_Fecha And Not Existe_Codigo Then sSQL = sSQL & "ORDER BY Fecha "
               If Not Existe_Fecha And Existe_Codigo Then sSQL = sSQL & "ORDER BY Codigo "
               Select_Adodc AdoCopyOrigen, sSQL
               With AdoCopyOrigen.Recordset
                If .RecordCount > 0 Then
                    Do While Not .EOF
                       sSQL = "SELECT * " _
                            & "FROM " & NombreTabla & " " _
                            & "WHERE Item = '" & NumEmpresa & "' "
                       For J = 0 To .Fields.Count - 1
                           If .Fields(J).Name <> "ID" Or .Fields(J).Name <> "Item" Then
                               Select Case .Fields(J).Type
                                 Case TadBoolean
                                      sSQL = sSQL & "AND " & .Fields(J).Name & " = " & CBool(.Fields(J)) & " "
                                 Case TadText, TadMemo
                                      sSQL = sSQL & "AND " & .Fields(J).Name & " = '" & .Fields(J) & "' "
                                 Case TadDate, TadDate1
                                      sSQL = sSQL & "AND " & .Fields(J).Name & " = #" & BuscarFecha(.Fields(J)) & "# "
                                 Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
                                      sSQL = sSQL & "AND " & .Fields(J).Name & " = " & .Fields(J) & " "
                               End Select
                           End If
                       Next J
                       sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                       MsgBox sSQL
                       Select_AdoDB RegInsertarDB, sSQL
                      'Insertamos el registro en la tabla destino
                       If RegInsertarDB.RecordCount <= 0 Then
                          SetAdoAddNew NombreTabla
                          For J = 0 To .Fields.Count - 1
                              Select Case .Fields(J).Name
                                Case "Item"
                                     SetAdoFields .Fields(J).Name, NumEmpresa
                                Case "ID"
                                    'No hace nada
                                Case Else
                                     SetAdoFields .Fields(J).Name, .Fields(J)
                              End Select
                          Next J
                          SetAdoUpdate
                       End If
                       RegInsertarDB.Close
                      .MoveNext
                    Loop
                End If
               End With
            End If
        Next i
        RatonNormal
        MsgBox "Proceso terminado con éxito"
        Unload FCopyEmpresa
     End If
  End If
End Sub

Private Sub Command2_Click()
  Unload FCopyEmpresa
End Sub

Private Sub Command3_Click()
   FrmEntidad.Visible = False
   DLEmpresa.SetFocus
End Sub

Private Sub Command4_Click()
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
         'Determinar que tipo de bases que utilizamos
          Select Case .Fields("Tipo_Base")
            Case "SQL SERVER"
                 AdoStrCnnBackup = "Data Source=" & .Fields("IP_VPN_RUTA") & ";" & vbCrLf _
                                 & "Initial Catalog=" & .Fields("Base_Datos") & ";" & vbCrLf _
                                 & "Provider=SQLOLEDB.1;" & vbCrLf _
                                 & "UID=" & .Fields("Usuario_DB") & ";" & vbCrLf _
                                 & "PWD=" & .Fields("Contraseña_DB") & ";"
                 SQL_Server = True
            Case "MY SQL"
                 AdoStrCnnBackup = "DRIVER={MySQL ODBC 3.51 Driver};" & vbCrLf _
                                 & "SERVER=" & .Fields("IP_VPN_RUTA") & ";" & vbCrLf _
                                 & "DATABASE=" & .Fields("Base_Datos") & ";" & vbCrLf _
                                 & "USER=" & .Fields("Usuario_DB") & ";" & vbCrLf _
                                 & "PASSWORD=" & .Fields("Contraseña_DB") & ";" & vbCrLf _
                                 & "PORT=" & .Fields("Puerto") & ";" & vbCrLf _
                                 & "OPTION=3;"
            Case "ACCESS"
                 AdoStrCnnBackup = "Data Source=" & .Fields("IP_VPN_DB") & "\" & .Fields("Base_Datos") & ".MDB;" & vbCrLf _
                                 & "Provider=Microsoft.Jet.OLEDB.4.0;" & vbCrLf _
                                 & "Persist Security Info=False;"
          End Select
         'MsgBox recuperar_IP(.Fields("IP_VPN_RUTA"))
          If Not Ping_PC(.Fields("IP_VPN_RUTA")) Then
             MsgBox "LA CONEXION NO ESTA ESTABLECIDA" & vbCrLf _
                  & "POR FAVOR LLAME AL ADMINISTRADOR" & vbCrLf _
                  & "PARA QUE CONECTE LA VPN"
             End
          Else
             Cadena = AdoStrCnn
             AdoStrCnn = AdoStrCnnBackup
             ConectarAdodc AdoCopyOrigen
             ConectarAdodc AdoEmp
             AdoStrCnn = Cadena
            '=======================================================================
             sSQL = "SELECT Empresa,Item " _
                  & "FROM Empresas " _
                  & "WHERE Item <> '" & Ninguno & "' " _
                  & "ORDER BY Empresa,Item "
             SelectDB_List DLEmpresa, AdoEmp, sSQL, "Empresa"
             RatonNormal
             If AdoEmp.Recordset.RecordCount <= 0 Then
                ConectarAdodc AdoCopyOrigen
                ConectarAdodc AdoEmp
                AdoStrCnnBackup = AdoStrCnn
                sSQL = "SELECT Empresa,Item " _
                     & "FROM Empresas " _
                     & "WHERE Item <> '" & NumEmpresa & "' " _
                     & "ORDER BY Empresa,Item "
                SelectDB_List DLEmpresa, AdoEmp, sSQL, "Empresa"
                RatonNormal
                MsgBox "No tiene empresas a quien copiar"
             End If
             DLEmpresa.SetFocus
          End If
       End If
   Else
      MsgBox "No hay Entidad asignada para copiar datos de origen"
   End If
  End With
  FrmEntidad.Visible = False
  DLEmpresa.SetFocus
End Sub

Private Sub DLEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And (KeyCode = vbKeyE) Then
     sSQL = "SELECT * " _
          & "FROM Empresas_Externas " _
          & "WHERE Entidad_Comercial <> '.' " _
          & "ORDER BY Entidad_Comercial "
     SelectDB_List DLEntidad, AdoEntidad, sSQL, "Entidad_Comercial"
     If AdoEntidad.Recordset.RecordCount > 0 Then
        FrmEntidad.Visible = True
        DLEntidad.SetFocus
     Else
        MsgBox "LLAME A SU PROVEEDOR PARA QUE CONFIGURE" & vbCrLf _
             & "ESTA OPCION Y PODER DISFRUTAR LA NUEVA" & vbCrLf _
             & "FORMA DE CONECTARCE CON OTRAS ENTIDADES" & vbCrLf _
             & "EMAIL: diskcover@msn.com" & vbCrLf _
             & "diskcoversystem@msn.com" & vbCrLf _
             & "TELEFONO BPX: 593-02-6052430" & vbCrLf
        DLEmpresa.SetFocus
     End If
  End If
End Sub

Private Sub Form_Activate()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim NombreTabla As String
Dim TablaCopiar As Boolean
Dim Conexion_Local As Tipo_Conexion
Dim Conexiones() As Tipo_Conexion

  RatonNormal
 'Crea variables de objeto para los objetos de acceso a datos.
 '=======================================================================
  LstTablas.Visible = False
  Leer_Datos_Conexion Conexion_Local
 'MsgBox AdoStrCnn
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  LstTablas.Clear
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
       'Llenamos la lista de Tablas
        TablaCopiar = True
        NombreTabla = RstSchema!TABLE_NAME
        If MidStrg(NombreTabla, 1, 7) = "Asiento" Then TablaCopiar = False
        If MidStrg(NombreTabla, 1, 7) = "Formato" Then TablaCopiar = False
        If MidStrg(NombreTabla, 1, 7) = "Balance" Then TablaCopiar = False
        If MidStrg(NombreTabla, 1, 5) = "Tabla" Then TablaCopiar = False
        If MidStrg(NombreTabla, 1, 4) = "Tipo" Then TablaCopiar = False
        If NombreTabla = "Codigos" Then TablaCopiar = False
        If NombreTabla = "Empresas" Then TablaCopiar = False
        If NombreTabla = "Modulos" Then TablaCopiar = False
        If NombreTabla = "Saldo_Diarios" Then TablaCopiar = False
       ' If NombreTabla = "Codigos" Then TablaCopiar = false
        If TablaCopiar Then LstTablas.AddItem NombreTabla
     End If
     RstSchema.MoveNext
  Loop
  For i = 0 To LstTablas.ListCount - 1
      LstTablas.Selected(i) = True
  Next i
  RstSchema.Close
  AdoCon1.Close
'=======================================================================
  LstTablas.Visible = True
  sSQL = "SELECT Empresa,Item " _
       & "FROM Empresas " _
       & "WHERE Item <> '" & NumEmpresa & "' " _
       & "ORDER BY Empresa,Item "
  SelectDB_List DLEmpresa, AdoEmp, sSQL, "Empresa"
  RatonNormal
  If AdoEmp.Recordset.RecordCount <= 0 Then
     MsgBox "No tiene empresas a quien copiar"
     Unload FCopyCat
  End If
End Sub

Private Sub Form_Load()
  CentrarForm FCopyEmpresa
  ConectarAdodc AdoCta
  ConectarAdodc AdoEmp
  ConectarAdodc AdoEntidad
  ConectarAdodc AdoCopyOrigen
  AdoStrCnnBackup = AdoStrCnn
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
  If KeyCode = vbKeyEscape Then Unload FCopyEmpresa
End Sub

Private Sub DLEntidad_KeyUp(KeyCode As Integer, Shift As Integer)
  With AdoEntidad.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Entidad_Comercial = '" & DLEntidad & "' ")
       If Not .EOF Then
          TxtReferencia = "IP VPN    : " & .Fields("IP_VPN_RUTA") & vbCrLf _
                        & "USUARIO   : " & .Fields("Usuario_PC") & vbCrLf _
                        & "CLAVE     : " & .Fields("Contraseña_PC") & vbCrLf _
                        & "BASE DATOS: " & .Fields("Base_Datos") & vbCrLf _
                        & "CLAVE DB  : " & .Fields("Contraseña_DB") & vbCrLf _
                        & "PUERTO    : " & .Fields("Puerto") & vbCrLf _
                        & "TEAMVIEWER: " & .Fields("ID_Conexion") & vbCrLf _
                        & "CLAVE TVW : " & .Fields("ID_Clave")
       End If
   End If
  End With
End Sub


