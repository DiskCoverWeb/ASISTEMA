VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form IngHabitacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reservaciones"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataList DLArt 
      Bindings        =   "IngHabit.frx":0000
      DataSource      =   "AdoArt"
      Height          =   1425
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   4640
      _ExtentX        =   8176
      _ExtentY        =   2514
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   330
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Art"
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
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Articulo"
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   240
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoAcomp 
      Height          =   330
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Acomp"
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
   Begin VB.TextBox TextCta 
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
      Left            =   2835
      MaxLength       =   12
      TabIndex        =   6
      Top             =   2310
      Width           =   1905
   End
   Begin VB.TextBox TextLinea 
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
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1995
      Width           =   3690
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
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
      Left            =   4830
      Picture         =   "IngHabit.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1155
      Width           =   1065
   End
   Begin VB.TextBox TextCodigo 
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
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1995
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Left            =   4830
      Picture         =   "IngHabit.frx":0457
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Width           =   1065
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
      Left            =   4830
      Picture         =   "IngHabit.frx":0899
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Label LabelCliente 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Habitacion Libre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   105
      TabIndex        =   11
      Top             =   2520
      Width           =   4635
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ESTADO DE LA HABITACION"
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
      TabIndex        =   10
      Top             =   2625
      Width           =   4635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DESCIPCION DE LA HABITACION"
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
      Left            =   1050
      TabIndex        =   2
      Top             =   1785
      Width           =   3690
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " VALOR DE LA HABITACION"
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
      TabIndex        =   5
      Top             =   2310
      Width           =   2745
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. Hab."
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
      TabIndex        =   1
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMBRE DEL PRODUCTO"
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
      TabIndex        =   0
      Top             =   105
      Width           =   4640
   End
End
Attribute VB_Name = "IngHabitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  GrabarLineas
End Sub

Private Sub Command2_Click()
  Unload IngHabitacion
End Sub

Private Sub Command3_Click()
  Codigo = TextCodigo.Text
  sSQL = "DELETE * FROM Habitaciones "
  sSQL = sSQL & "WHERE No_Hab ='" & Codigo & "' "
  ConectarAdoExecute sSQL
  sSQL = "SELECT (No_Hab & Space(5) & Descripcion) As CodHabit " _
        & "FROM Habitaciones " _
        & "ORDER BY No_Hab "
  SelectDBList DLArt, AdoArt, sSQL, "CodHabit"
End Sub

Private Sub DLArt_DblClick()
  SiguienteControl
End Sub

Private Sub DLArt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DLArt_LostFocus()
  Codigo = SinEspaciosIzq(DLArt.Text)
  If Codigo = "" Then Codigo = Ninguno
  LlenarLineas Codigo
End Sub

Private Sub Form_Activate()
   sSQL = "SELECT (No_Hab & Space(5) & Descripcion) As CodHabit " _
        & "FROM Habitaciones " _
        & "ORDER BY No_Hab "
   SelectDBList DLArt, AdoArt, sSQL, "CodHabit"
   RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm IngHabitacion
   ConectarAdodc AdoArt
   ConectarAdodc AdoAcomp
   ConectarAdodc AdoCliente
   ConectarAdodc AdoArticulo
End Sub

Private Sub TextCodigo_LostFocus()
  TextoValido TextCodigo, , True
End Sub

Private Sub TextCta_LostFocus()
  TextoValido TextCta
End Sub

Private Sub TextLinea_LostFocus()
  TextoValido TextLinea, , True
End Sub

Public Sub LlenarLineas(CodigoArt As String)
  LabelCliente.Caption = " Habitación Libre"
  sSQL = "SELECT * FROM Habitaciones "
  sSQL = sSQL & "WHERE No_Hab ='" & CodigoArt & "' "
  SelectData AdoArticulo, sSQL, False
  With AdoArticulo.Recordset
   If .RecordCount > 0 Then
       TextCodigo.Text = .Fields("No_Hab")
       TextLinea.Text = .Fields("Descripcion")
       TextCta.Text = .Fields("Valor_Hab")
       CodigoCli = .Fields("CodigoC")
       sSQL = "SELECT * FROM Clientes "
       sSQL = sSQL & "WHERE Codigo ='" & CodigoCli & "' "
       SelectData AdoCliente, sSQL, False
       If AdoCliente.Recordset.RecordCount > 0 Then
          Cadena = AdoCliente.Recordset.Fields("Apellidos")
          Cadena = Cadena & " " & AdoCliente.Recordset.Fields("Nombres")
          LabelCliente.Caption = Cadena
       End If
       sSQL = "SELECT No_Hab,Acompañante "
       sSQL = sSQL & "FROM Acompañantes "
       sSQL = sSQL & "WHERE No_Hab ='" & TextCodigo.Text & "' "
       sSQL = sSQL & "AND Factura_No = 0 "
       SelectDataGrid DGAcomp, AdoAcomp, sSQL
   Else
       MsgBox "Esta Habitacion no exite."
   End If
  End With
End Sub

Public Sub GrabarLineas()
  Codigo = TextCodigo.Text
  Mensajes = "Esta seguro de Grabar: " _
           & TextCodigo.Text & "."
  Titulo = "Pregunta de grabación"
  If BoxMensaje = 6 Then
     sSQL = "SELECT * FROM Habitaciones "
     sSQL = sSQL & "WHERE No_Hab = '" & Codigo & "' "
     SelectData AdoArticulo, sSQL, False
     With AdoArticulo.Recordset
          If .RecordCount > 0 Then
             '.Edit
          Else
             .AddNew
             .Fields("No_Hab") = TextCodigo.Text
             .Fields("Ocupada") = False
             .Fields("Ingreso") = FechaSistema
             .Fields("Salida") = FechaSistema
             .Fields("CodigoC") = Ninguno
          End If
         .Fields("Descripcion") = TextLinea.Text
         .Fields("Valor_Hab") = TextCta.Text
         .Update
          sSQL = "SELECT (No_Hab & Space(5) & Descripcion) As CodHabit " _
               & "FROM Habitaciones " _
               & "ORDER BY No_Hab "
          SelectDBList DLArt, AdoArt, sSQL, "CodHabit"
     End With
  End If
  Nuevo = False
End Sub
