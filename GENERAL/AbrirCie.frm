VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form AbrirPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REABRIR PERIODOS"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   Icon            =   "AbrirCie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstPeriodo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "AbrirCie.frx":030A
      Left            =   105
      List            =   "AbrirCie.frx":030C
      TabIndex        =   1
      Top             =   420
      Width           =   2010
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
      Left            =   2205
      Picture         =   "AbrirCie.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1155
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Re-Abrir Periodo"
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
      Left            =   2205
      Picture         =   "AbrirCie.frx":0618
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1800
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
      Caption         =   "Seleccione el Periodo"
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
      Width           =   2010
   End
End
Attribute VB_Name = "AbrirPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IdTime As Long
Dim strCnn As String
Dim itmX As ListItem
  RatonNormal
  Mensajes = "Al querer reabrir el período seleccionado al " & LstPeriodo.Text & "," & vbCrLf & vbCrLf _
           & "se eliminará el asiento inicial del periodo actual, y se" & vbCrLf & vbCrLf _
           & "procederá a restaurar los  datos  del  cierre  anterior," & vbCrLf & vbCrLf _
           & "perdiendo la información procesada del Periodo actual." & vbCrLf & vbCrLf & vbCrLf _
           & "ESTA SEGURO DE REABRIR EL PERIODO AL " & LstPeriodo.Text & "?"
  Titulo = "REAPERTURA DEL PERIODO FISCAL AL:" & LstPeriodo.Text
  If BoxMensaje = vbYes Then
     RatonReloj
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Mensaje_Box = "Reabriendo Periodo"
     Progreso_Iniciar
      Periodo_Contable = LstPeriodo.Text
      If Periodo_Contable = "Periodo Actual" Then Periodo_Contable = Ninguno
      Contador = 0
      LstPeriodo.Clear
    ' Crea variables de objeto para los objetos de acceso a datos.
      Set AdoCon1 = New ADODB.Connection
      AdoCon1.open AdoStrCnn
      Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
      Do Until RstSchema.EOF
         If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
           'Llenamos la lista de Tablas
            LstPeriodo.AddItem RstSchema!TABLE_NAME
            Contador = Contador + 1
         End If
         RstSchema.MoveNext
      Loop
      Progreso_Barra.Valor_Maximo = Contador + 10
      Progreso_Barra.Mensaje_Box = "Eliminando Asiento inicial"
      Progreso_Esperar
      
      sSQL = "DELETE " _
           & "FROM Transacciones " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '.' " _
           & "AND TP = 'CD' " _
           & "AND Numero = 1 "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE " _
           & "FROM Trans_SubCtas " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '.' " _
           & "AND TP = 'CD' " _
           & "AND Numero = 1 "
      Ejecutar_SQL_SP sSQL
      sSQL = "DELETE " _
           & "FROM Comprobantes " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '.' " _
           & "AND TP = 'CD' " _
           & "AND Numero = 1 "
      Ejecutar_SQL_SP sSQL
      For i = 0 To LstPeriodo.ListCount - 1
          Progreso_Barra.Mensaje_Box = "RP: " & LstPeriodo.List(i)
          Progreso_Esperar
          If MidStrg(LstPeriodo.List(i), 1, 7) <> "Asiento" And MidStrg(LstPeriodo.List(i), 1, 8) <> "Balances" And _
             LstPeriodo.List(i) <> "Seteos_Documentos" And LstPeriodo.List(i) <> "Formato" And _
             LstPeriodo.List(i) <> "Clientes" Then
            'MsgBox "RP: " & LstPeriodo.List(I) & "..."
             Si_No = False
             Modificar = False
             sSQL = "SELECT TOP(1) * " _
                  & "FROM " & LstPeriodo.List(i) & " "
             Select_Adodc AdoPeriodo, sSQL
             RatonReloj
             With AdoPeriodo.Recordset
              For J = 0 To .Fields.Count - 1
                  If .Fields(J).Name = "Item" Then Si_No = True
                  If .Fields(J).Name = "Periodo" Then Modificar = True
              Next J
             End With
             If Modificar And Si_No Then
                If MidStrg(LstPeriodo.List(i), 1, 8) = "Catalogo" Then
                   sSQL = "DELETE " _
                        & "FROM " & LstPeriodo.List(i) & " " _
                        & "WHERE Item = '" & NumEmpresa & "' " _
                        & "AND Periodo = '.' "
                   Ejecutar_SQL_SP sSQL
                End If
                sSQL = "UPDATE " & LstPeriodo.List(i) & " " _
                     & "SET Periodo = '.' " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' "
                Ejecutar_SQL_SP sSQL
             End If
             RatonNormal
          End If
      Next i
      RatonNormal
     Progreso_Final
     Periodo_Contable = Ninguno
     MsgBox "Proceso Terminado"
  End If
  Unload AbrirPeriodo
End Sub

Private Sub Command2_Click()
  Unload AbrirPeriodo
End Sub

Private Sub Form_Activate()
  RatonReloj
  LstPeriodo.Clear
  LstPeriodo.AddItem "Periodo Actual"
  sSQL = "SELECT Periodo " _
       & "FROM Comprobantes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo <> '.' " _
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
  CentrarForm AbrirPeriodo
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
     Periodo_Contable = InputBox("Digite el Periodo a Reabrir:", "CAMBIOS DE PERIODOS", ".")
     LstPeriodo.AddItem Periodo_Contable
  End If
End Sub
