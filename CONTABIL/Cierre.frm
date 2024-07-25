VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form Cierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Periodo"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
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
      Left            =   2835
      Picture         =   "Cierre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1050
      Width           =   960
   End
   Begin VB.ListBox LstMeses 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   420
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc AdoCierre 
      Height          =   330
      Left            =   210
      Top             =   2730
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
      Caption         =   "Cierre"
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
      Left            =   2835
      Picture         =   "Cierre.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo Contable"
      BeginProperty Font 
         Name            =   "Courier New"
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
      Width           =   2535
   End
End
Attribute VB_Name = "Cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim NoMes As Integer
Dim Anio As String
Dim Mes As String
Dim Mes_C As Integer
Dim UnaVez As Boolean
Dim Fecha_Mes As String
  UnaVez = True
  FechaCierre = FechaSistema
  For NoMes = 0 To LstMeses.ListCount - 1
      Anio = SinEspaciosIzq(LstMeses.List(NoMes))
      Mes = SinEspaciosDer(LstMeses.List(NoMes))
      If LstMeses.Selected(NoMes) Then Mes_C = 1 Else Mes_C = 0
      Fecha_Mes = "01/" & LetrasMeses(Mes) & "/" & Anio
      If LstMeses.Selected(NoMes) = False And UnaVez Then
         FechaCierre = Fecha_Mes
         UnaVez = False
      End If
      sSQL = "UPDATE Fechas_Balance " _
           & "SET Cerrado = " & Mes_C & " " _
           & "WHERE MidStrg(Detalle,1,4) = '" & Anio & "' " _
           & "AND Fecha_Inicial = #" & BuscarFecha(Fecha_Mes) & "# " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQL
  Next NoMes
  Control_Procesos Normal, "Cambio periodo mensual al: " & FechaCierre
  Unload Cierre
End Sub

Private Sub Command2_Click()
  Unload Cierre
End Sub

Private Sub Form_Activate()
Dim NoMes As Byte
Dim IMes As Integer
  LstMeses.Clear
  IMes = 0
  sSQL = "SELECT * " _
       & "FROM Fechas_Balance " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND ISNUMERIC(MidStrg(Detalle,1,4)) <> " & adFalse & " " _
       & "ORDER BY Fecha_Inicial "
  Select_Adodc AdoCierre, sSQL
  With AdoCierre.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          LstMeses.AddItem .fields("Detalle"), IMes
          LstMeses.Selected(IMes) = CBool(.fields("Cerrado"))
          IMes = IMes + 1
         .MoveNext
       Loop
   End If
  End With
  LstMeses.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Cierre
  ConectarAdodc AdoCierre
  FechaCierre = FechaSistema
End Sub

Private Sub LstMeses_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Nuevo_Anio As String
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload Me
  If CtrlDown And KeyCode = vbKeyA Then
     Nuevo_Anio = InputBox("INGRESE EL AÑO A PROCESAR:", "AÑO DE PROCESO", Year(FechaSistema))
     Crear_Cierre_Mes Nuevo_Anio
     Unload Me
  End If
End Sub
