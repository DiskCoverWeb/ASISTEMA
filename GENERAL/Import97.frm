VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Importar97 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IMPORTACION DE DATOS"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc AdoNomina 
      Height          =   330
      Left            =   105
      Top             =   840
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
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
      Caption         =   "Nomina"
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
   Begin VB.CommandButton Command6 
      Caption         =   "&Nomina 97"
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
      Picture         =   "Import97.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3990
      Picture         =   "Import97.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   1470
      TabIndex        =   3
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoNomina97 
      Height          =   330
      Left            =   2415
      Top             =   840
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Sistema\EMPRESA\CETAS\nomina.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Sistema\EMPRESA\CETAS\nomina.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Nomina97"
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
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
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
      Left            =   1470
      TabIndex        =   2
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
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
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "Importar97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Command6_Click()
Dim PathTemp1 As String
Dim PathTemp2 As String
Dim NumFile As Long
  ConSubDir = True
  FUnidad.Show 1
  If RutaSubDirTemp <> "" Then
  PathTemp1 = PathEmpresa
  PathTemp2 = AdoStrCnn
  RatonReloj
' Buscamos la cadena de conección a la base
  RutaGeneraFile = RutaSistema & "\CONECTAR.TXT"
  NumFile = FreeFile
  AdoStrCnn = ""
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
       AdoStrCnn = AdoStrCnn & Input(1, #NumFile) ' Obtiene un carácter.
    Loop
  Close #NumFile
  PathEmpresa = UCase(RutaSubDirTemp & "\NOMINA.MDB")
  AdoStrCnn = AdoStrCnn & "Data Source = " & PathEmpresa
  'MsgBox AdoStrCnn
  ConectarAdodc AdoNomina97
  PathEmpresa = PathTemp1
  AdoStrCnn = PathTemp2
  
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "DELETE * " _
       & "FROM Trans_RolHoras " _
       & "WHERE Fecha " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  ConectarAdoExecute sSQL
'''  sSQL = "SELECT * " _
'''       & "FROM Trans_RolHoras " _
'''       & "WHERE Fecha " _
'''       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  SelectAdodc AdoNomina, sSQL
  Si_No = SQL_Server
  SQL_Server = False
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "SELECT * " _
       & "FROM [horas trabajadas] " _
       & "WHERE Fecha " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "ORDER BY códigoso,fecha "
  'MsgBox sSQL
  SelectAdodc AdoNomina97, sSQL
  SQL_Server = Si_No
  With AdoNomina97.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Do While Not .EOF
          Importar97.Caption = "IMPORTACION DE DATOS: " & Format(.Fields("códigoso"), "0000000000")
          Total = .Fields("horatrabajada") * .Fields("Valorhora")
          'MsgBox .Fields("fecha")
          SetAdoAddNew "Trans_RolHoras"
          SetAdoFields "Codigo", Format(.Fields("códigoso"), "0000000000")
          SetAdoFields "Fecha", .Fields("fecha")
          SetAdoFields "Horas", Round(.Fields("horatrabajada"), 2)
          SetAdoFields "Ing_Liquido", Round(Total, 2)
          SetAdoFields "Valor_Hora", Round(.Fields("Valorhora"), 2)
          SetAdoFields "Vacaciones", Round(.Fields("Vacaciones"), 2)
          SetAdoFields "Fondo_Reserva", Round(.Fields("Fondo"), 2)
          SetAdoFields "IESS_Per", Round(Total * .Fields("IEESS") / 100, 2)
          SetAdoFields "IESS_Pat", Round(.Fields("tiesspatronal"), 2)
          SetAdoFields "Certificado", Round(Total * 0.08, 2)
          SetAdoFields "Aporte_Adm", Round(Total * 0.02, 2)
          SetAdoFields "Item", NumEmpresa
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  MsgBox "Importacion finalizada con éxito"
  Else
    RatonNormal
    MsgBox "No hay datos importados"
  End If
  Unload Me
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Importar97
  ConectarAdodc AdoNomina
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

