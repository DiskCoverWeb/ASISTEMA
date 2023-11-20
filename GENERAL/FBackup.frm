VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FBackup 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "REGISTRO DE TABLAS DE ORIGEN"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   7470
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   750
      Left            =   5775
      TabIndex        =   7
      Top             =   1050
      Width           =   1590
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Realizar Respaldo en linea"
      Height          =   855
      Left            =   5775
      TabIndex        =   6
      Top             =   105
      Width           =   1590
   End
   Begin VB.ListBox LstTablas 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1110
      Left            =   105
      TabIndex        =   4
      Top             =   735
      Width           =   5580
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2835
      TabIndex        =   10
      Top             =   1470
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox PictTotal 
      Height          =   330
      Left            =   105
      ScaleHeight     =   270
      ScaleWidth      =   7200
      TabIndex        =   5
      Top             =   1890
      Width           =   7260
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   105
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   2835
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2640
   End
   Begin MSAdodcLib.Adodc AdoBases 
      Height          =   330
      Left            =   105
      Top             =   2205
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Bases"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoBackup 
      Height          =   330
      Left            =   2625
      Top             =   2205
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Backup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   4410
      TabIndex        =   2
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
      Left            =   3150
      TabIndex        =   1
      Top             =   105
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fechas de Respaldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
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
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATOS A RESPALDAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   420
      Width           =   5580
   End
End
Attribute VB_Name = "FBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'carla.jativa@dhl.com

Option Explicit

Public Function ConverType(CampoType As Integer) As String
Dim TipoCampo As String
  If SQL_Server Then
     TipoCampo = SinEspaciosIzq(TablaNew(CampoType).TipoSQL)
  Else
     TipoCampo = SinEspaciosIzq(TablaNew(CampoType).TipoAccess)
  End If
  Select Case TipoCampo
    Case "BIT", "TINYINT", "SMALLINT", "INT", "BYTE", "SHORT", "LONG", "INTEGER"
         ConverType = "CInt(" & TablaNew(CampoType).Campo & ")"
    Case "REAL", "FLOAT", "MONEY", "DECIMAL", "SINGLE", "DOUBLE", "CURRENCY"
         ConverType = "CSng(" & TablaNew(CampoType).Campo & ")"
    Case "NTEXT", "NVARCHAR", "TEXT", "LONGTEXT"
         ConverType = "CStr(" & TablaNew(CampoType).Campo & ")"
    Case Else
         ConverType = TablaNew(CampoType).Campo
  End Select
End Function

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub CommandButton2_Click()
Dim Si_Actualiza As Boolean
Dim Nombre_Key As String
Dim AdoStrCnnTemp As String
Dim Nombre_Tabla As String
Dim Existe_Fecha As Boolean
Dim Existe_Item As Boolean
Dim Existe_Periodo As Boolean
Dim Existe_Where As Boolean
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  
  AdoStrCnnTemp = AdoStrCnn
  AdoStrCnn = AdoStrCnnBackup
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  Progreso_Barra.Mensaje_Box = "PROGRESO DEL RESPALDO"
  Progreso_Esperar
  UPD_Listar_Tablas LstTablas
  UPD_Actualizar Dir1, File1, LstTablas, True
  AdoStrCnn = AdoStrCnnTemp
  ConectarAdodc AdoBases
  ConectarAdodcBackup AdoBackup
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = LstTablas.ListCount
  Progreso_Barra.Mensaje_Box = "PROGRESO DEL RESPALDO"
  Progreso_Esperar

  For I = 0 To LstTablas.ListCount - 1
     'Averiguamos si la tabla esta vacia entonces la eliminamos
      Nombre_Tabla = LstTablas.List(I)
      Progreso_Barra.Mensaje_Box = "Respaldando: " & Nombre_Tabla
      Progreso_Esperar
      Existe_Fecha = False
      Existe_Item = False
      Existe_Where = False
      Existe_Periodo = False
      SQL1 = ""
      sSQL = "SELECT TOP 1 * " _
           & "FROM [" & Nombre_Tabla & "] "
      SelectAdodc AdoBases, sSQL
      With AdoBases.Recordset
         For K = 0 To .Fields.Count - 1
             If .Fields(K).Name = "Item" Then Existe_Item = True
             If .Fields(K).Name = "Fecha" Then Existe_Fecha = True
             If .Fields(K).Name = "Periodo" Then Existe_Periodo = True
         Next K
      End With
      If Existe_Item Then
         SQL1 = SQL1 & "WHERE Item = '" & NumEmpresa & "' "
         Existe_Where = True
      End If
      If Existe_Periodo Then
         If Existe_Where Then
            SQL1 = SQL1 & "AND Periodo = '" & Periodo_Contable & "' "
         Else
            SQL1 = SQL1 & "WHERE Periodo = '" & Periodo_Contable & "' "
            Existe_Where = True
         End If
      End If
      If Existe_Fecha Then
         If Existe_Where Then
            SQL1 = SQL1 _
                 & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
         Else
            SQL1 = SQL1 _
                 & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
         End If
      End If
      sSQL = "DELETE * " _
           & "FROM [" & Nombre_Tabla & "] " _
           & SQL1
      ConectarAdoExecuteBackup sSQL
      sSQL = "SELECT * " _
           & "FROM [" & Nombre_Tabla & "] " _
           & SQL1
      SelectAdodc AdoBackup, sSQL
      
      sSQL = "SELECT * " _
           & "FROM [" & Nombre_Tabla & "] " _
           & SQL1
      If Existe_Fecha Then sSQL = sSQL & "ORDER BY Fecha "
      SelectAdodc AdoBases, sSQL
      With AdoBases.Recordset
       If .RecordCount > 0 Then
           Do While Not .EOF
              For J = 0 To .Fields.Count - 1
                  SetAddNew AdoBackup
                  SetFields AdoBackup, .Fields(J).Name, .Fields(J)
                  SetUpdate AdoBackup
              Next J
             .MoveNext
           Loop
       End If
      End With
  Next I
  RatonNormal
  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
End Sub

Private Sub Form_Activate()
Dim CarBase As String
  RatonNormal
  If Hacer_Backup Then
    'Buscamos la cadena de conección a la base
     RutaGeneraFile = RutaSistema & "\Backup.txt"
     AdoStrCnnBackup = ""
     NumFile = FreeFile
     Open RutaGeneraFile For Input As #NumFile
     Do While Not EOF(NumFile)
        CarBase = Input(1, #NumFile) ' Obtiene un carácter.
        AdoStrCnnBackup = AdoStrCnnBackup & CarBase
     Loop
     Close #NumFile
     CommandButton2.SetFocus
  Else
     MsgBox "No se puede hacer respaldo en linea, pues no esta activada esta opcion"
     Unload Me
  End If
End Sub

Private Sub Form_Load()
  CentrarForm FBackup
  Dir1.Path = RutaEmpresa
  File1.Filename = Dir1.Path & "\*.MDB"
End Sub
