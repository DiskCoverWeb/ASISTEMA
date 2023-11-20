VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FMigracion 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MIGRACION DE BASES DE SQL SERVER A MY SQL"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CPeriodo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   4
      Top             =   525
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox CheqPeriodo 
      BackColor       =   &H00808080&
      Caption         =   "Por &Periodo"
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
      TabIndex        =   3
      Top             =   525
      Width           =   1590
   End
   Begin VB.ComboBox CEmpresa 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   1260
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Salir"
      Height          =   330
      Left            =   3885
      TabIndex        =   9
      Top             =   7245
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Migrar Base >"
      Height          =   330
      Left            =   5250
      TabIndex        =   10
      Top             =   7245
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3150
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12632256
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
   Begin VB.CheckBox CheqItem 
      BackColor       =   &H00808080&
      Caption         =   "Por &Empresa"
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
      TabIndex        =   5
      Top             =   945
      Width           =   1590
   End
   Begin VB.CheckBox CheqFechas 
      BackColor       =   &H00808080&
      Caption         =   "Por &Fechas:"
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
      Width           =   1590
   End
   Begin VB.ListBox LstTablas 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5100
      Left            =   105
      TabIndex        =   8
      Top             =   1995
      Width           =   6420
   End
   Begin MSAdodcLib.Adodc AdoPeriodo 
      Height          =   330
      Left            =   105
      Top             =   6405
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4515
      Top             =   6405
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   330
      Left            =   2415
      Top             =   6405
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
      Caption         =   "Empresa"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   1785
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12632256
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TABLAS A MIGRAR"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   6420
   End
End
Attribute VB_Name = "FMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AdoStrCnnMySQL As String
Dim RutaGeneraFile As String
Dim NameFile As String
Dim IdTabla As Long

Private Sub CheqFechas_Click()
   If CheqFechas.value = 1 Then
      MBFechaI.Visible = True
      MBFechaF.Visible = True
   Else
      MBFechaI.Visible = False
      MBFechaF.Visible = False
   End If
End Sub

Private Sub CheqItem_Click()
   If CheqItem.value = 1 Then CEmpresa.Visible = True Else CEmpresa.Visible = False
End Sub

Private Sub CheqPeriodo_Click()
   If CheqPeriodo.value = 1 Then CPeriodo.Visible = True Else CPeriodo.Visible = False
End Sub

Private Sub Command1_Click()
Dim NombreTabla As String
Dim BID As Boolean
Dim BItem As Boolean
Dim BFecha As Boolean
Dim BPeriodo As Boolean
Dim EsWhere As Boolean
Dim vItem As String
Dim AdoStrCnnTemp As String

'AdoStrCnn
'AdoStrCnn1
    Progreso_Iniciar
    Progreso_Barra.Mensaje_Box = "Determinando Tablas y Campos"
    Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + LstTablas.ListCount
    Progreso_Esperar
    vItem = SinEspaciosIzq(CEmpresa)
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    For IdTabla = 0 To LstTablas.ListCount - 1
        LstTablas.Text = LstTablas.List(IdTabla)
        LstTablas.Refresh
        
        BID = False
        BItem = False
        BFecha = False
        BPeriodo = False
        EsWhere = False
        If InStr(LstTablas.List(IdTabla), "ID") > 0 Then BID = True
        If InStr(LstTablas.List(IdTabla), "Item") > 0 Then BItem = True
        If InStr(LstTablas.List(IdTabla), "Fecha ") > 0 Then BFecha = True
        If InStr(LstTablas.List(IdTabla), "Periodo") > 0 Then BPeriodo = True
        
        NombreTabla = SinEspaciosIzq(LstTablas.List(IdTabla))
        Progreso_Barra.Mensaje_Box = "Tabla: " & NombreTabla
        Progreso_Esperar
        sSQL = "DELETE " _
             & "FROM " & NombreTabla & " "
        If BPeriodo Then
           If Not EsWhere Then sSQL = sSQL & "WHERE Periodo = '" & CPeriodo & "' " Else sSQL = sSQL & "AND Periodo = '" & CPeriodo & "' "
           EsWhere = True
        End If
        If BItem Then
           If Not EsWhere Then sSQL = sSQL & "WHERE Item = '" & vItem & "' " Else sSQL = sSQL & "AND Item = '" & vItem & "' "
           EsWhere = True
        End If
        If BFecha Then
           If Not EsWhere Then
              sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
           Else
              sSQL = sSQL & "AND Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
           End If
           EsWhere = True
        End If
        AdoStrCnnTemp = AdoStrCnn
        AdoStrCnn = AdoStrCnnMySQL
        Ejecutar_SQL_SP sSQL
        AdoStrCnn = AdoStrCnnTemp
        sSQL = Replace(sSQL, "DELETE ", "SELECT * ")
        If BFecha Then sSQL = sSQL & "ORDER BY Fecha "
        Select_Adodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Contador = 0
             Do While Not .EOF
                Progreso_Barra.Mensaje_Box = "Tabla: " & NombreTabla & " [" & Format(Contador / .RecordCount, "00%") & "]"
                Progreso_Esperar True
                SetAdoAddNew NombreTabla, True
                For i = 0 To .Fields.Count - 1
                    SetAdoFields .Fields(i).Name, .Fields(i)
                Next i
                SetAdoUpdateMySQL AdoStrCnnMySQL
                Contador = Contador + 1
                'MsgBox "...."
               .MoveNext
             Loop
         End If
        End With
        'FMigracion.Caption = "MIGRACION DE BASES DE SQL SERVER A MY SQL"
        'FMigracion.Refresh
    Next IdTabla
    MsgBox "Proceso Terminado Exitosamente"
    Unload FMigracion
End Sub

Private Sub Command4_Click()
  Unload FMigracion
End Sub

Private Sub Form_Activate()
Dim AdoCon1 As ADODB.Connection
Dim AdoDBCampos As ADODB.Recordset
Dim RstSchema As ADODB.Recordset
Dim NombreTabla As String
Dim VDetalle As String
Dim NID As String
Dim NItem As String
Dim NFecha As String
Dim NPeriodo As String
Dim Idx As Long

  RatonReloj
  LstTablas.Visible = False
  AdoStrCnnMySQL = ""
  NameFile = ""
  Cadena = Dir(RutaSistema & "\", vbNormal) 'Recupera la primera entrada.
  Do While Cadena <> ""
     If Cadena <> "." And Cadena <> ".." Then
        If (GetAttr(RutaSistema & "\" & Cadena) And vbNormal) = vbNormal Then
           If UCaseStrg(Cadena) = "MIGRACION.INI" Then NameFile = "Migracion.ini"
        End If
     End If
     Cadena = Dir
  Loop
  
  If Len(NameFile) > 1 Then
     RutaGeneraFile = RutaSistema & "\" & NameFile
     NumFile = FreeFile
     Open RutaGeneraFile For Input As #NumFile
     Do While Not EOF(NumFile)
        AdoStrCnnMySQL = AdoStrCnnMySQL & Input(1, #NumFile)   ' Obtiene un carácter.
     Loop
     Close #NumFile
       
    'Consultamos las cuentas de la tabla
     Set AdoCon1 = New ADODB.Connection
     AdoCon1.open AdoStrCnn
     Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
     Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And Mid(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
          NombreTabla = RstSchema!TABLE_NAME
          NID = ""
          NItem = ""
          NFecha = ""
          NPeriodo = ""
          sSQL = "SELECT TOP 1 * " _
               & "FROM " & NombreTabla & " "
          Select_AdoDB AdoDBCampos, sSQL
          If AdoDBCampos.RecordCount >= 0 Then
             For Idx = 0 To AdoDBCampos.Fields.Count - 1
                 If AdoDBCampos.Fields(Idx).Name = "ID" Then NID = "ID"
                 If AdoDBCampos.Fields(Idx).Name = "Item" Then NItem = "Item"
                 If AdoDBCampos.Fields(Idx).Name = "Fecha" Then NFecha = "Fecha"
                 If AdoDBCampos.Fields(Idx).Name = "Periodo" Then NPeriodo = "Periodo"
             Next Idx
          End If
          AdoDBCampos.Close
         'MsgBox NombreTabla & String(60 - Len(NombreTabla), " ") & NItem & " " & NFecha & " " & NPeriodo
          VDetalle = NombreTabla & String(30 - Len(NombreTabla), " ")
          If NItem <> "" Then VDetalle = VDetalle & " " & NItem
          If NFecha <> "" Then VDetalle = VDetalle & " " & NFecha
          If NPeriodo <> "" Then VDetalle = VDetalle & " " & NPeriodo
          LstTablas.AddItem VDetalle
       End If
       RstSchema.MoveNext
     Loop
     RstSchema.Close
    
    CEmpresa.Clear
    sSQL = "SELECT Empresa, Item " _
         & "FROM Empresas " _
         & "WHERE Item <> '_' " _
         & "ORDER BY Empresa, Item "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            CEmpresa.AddItem .Fields("Item") & " " & .Fields("Empresa")
           .MoveNext
         Loop
         CEmpresa.Text = CEmpresa.List(0)
     End If
    End With
    
    CPeriodo.Clear
    sSQL = "SELECT Periodo " _
         & "FROM Comprobantes " _
         & "WHERE Periodo <> '_' " _
         & "GROUP BY Periodo " _
         & "ORDER BY Periodo "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            CPeriodo.AddItem .Fields("Periodo")
           .MoveNext
         Loop
         CPeriodo.Text = CPeriodo.List(0)
     End If
    End With
  End If
  RatonNormal
  LstTablas.Visible = True
  FMigracion.Visible = True
  If Len(AdoStrCnnMySQL) < 1 Then
     MsgBox "No existe configuracion para migrar"
     Unload FMigracion
  End If
End Sub

Private Sub Form_Load()
  FMigracion.Visible = False
  CentrarForm FMigracion
  ConectarAdodc AdoAux
  ConectarAdodc AdoEmpresa
  ConectarAdodc AdoPeriodo
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

