VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FDisciplinaQuimestre 
   Caption         =   "RESUMEN DE NOTAS"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10965
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OpcT 
      Caption         =   "Los Dos Quimestres"
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
      Left            =   4305
      TabIndex        =   8
      Top             =   315
      Width           =   2115
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Consultar"
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
      Left            =   7455
      Picture         =   "FDiscQui.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
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
      Left            =   8610
      Picture         =   "FDiscQui.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Height          =   855
      Left            =   9765
      Picture         =   "FDiscQui.frx":1384
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   1065
   End
   Begin VB.OptionButton OpcPQBim2 
      Caption         =   "Segundo Periodo"
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
      Width           =   1800
   End
   Begin VB.OptionButton OpcPQBim1 
      Caption         =   "Primer Periodo"
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
      TabIndex        =   2
      Top             =   315
      Value           =   -1  'True
      Width           =   1590
   End
   Begin VB.OptionButton OpcPQ 
      Caption         =   "Primer Quimestre"
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
      TabIndex        =   4
      Top             =   735
      Width           =   1800
   End
   Begin VB.OptionButton OpcSQBim2 
      Caption         =   "Segundo Periodo"
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
      Left            =   2205
      TabIndex        =   6
      Top             =   525
      Width           =   1800
   End
   Begin VB.OptionButton OpcSQBim1 
      Caption         =   "Primer Periodo"
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
      Left            =   2205
      TabIndex        =   5
      Top             =   315
      Width           =   1590
   End
   Begin VB.OptionButton OpcSQ 
      Caption         =   "Segundo Quimestre"
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
      Left            =   2205
      TabIndex        =   7
      Top             =   735
      Width           =   2010
   End
   Begin VB.Frame Frame2 
      Caption         =   "LISTA DE NOTAS"
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
      Left            =   105
      TabIndex        =   9
      Top             =   1050
      Width           =   10725
      Begin VB.OptionButton OpcProm 
         Caption         =   "Promediales"
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
         Left            =   1890
         TabIndex        =   11
         Top             =   210
         Width           =   1485
      End
      Begin VB.OptionButton OpcNotas 
         Caption         =   "Por Curso"
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
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
      Begin MSDataListLib.DataList DLCurso 
         Bindings        =   "FDiscQui.frx":1C4E
         DataSource      =   "AdoAux"
         Height          =   300
         Left            =   3465
         TabIndex        =   12
         Top             =   210
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid DGResumen 
      Bindings        =   "FDiscQui.frx":1C63
      Height          =   5055
      Left            =   105
      TabIndex        =   16
      Top             =   1785
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   630
      Top             =   2730
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
   Begin MSAdodcLib.Adodc AdoDisciplina 
      Height          =   330
      Left            =   630
      Top             =   3045
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
      Caption         =   "Disciplina"
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
   Begin MSAdodcLib.Adodc AdoResumen 
      Height          =   330
      Left            =   4305
      Top             =   630
      Width           =   3060
      _ExtentX        =   5398
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
      Caption         =   "Resumen"
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
      Caption         =   "Primer Quimestre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Segundo Quimestre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "FDisciplinaQuimestre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Column, Row As Integer
Dim Index1, Index2, Index3, Index4 As Integer

Dim SumaHoriz(100) As Single
Dim VectNota(100) As Currency
Dim VectMate(100) As String
Dim VectCodMat(100) As String

Dim ContHoriz As Integer
Dim ContVert As Integer
Dim ContNotas As Integer

Private Sub Command1_Click()
  DGResumen.Visible = False
  CodigoInv = DLCurso.Text
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       If TipoDoc = "1" Then
          CodigoBenef = .Fields("Director")
          CodigoCorresp = .Fields("Secretario1")
       End If
       If TipoDoc = "2" Then
          CodigoBenef = .Fields("Rector")
          CodigoCorresp = .Fields("Secretario2")
       End If
       Codigo2 = .Fields("Anio_Lectivo")
   End If
  End With
  'MsgBox TipoDoc
  sSQL = "SELECT (CodigoE & ' - ' & Detalle) As Cursos,Seccion,Detalle As Paralelo,CodigoE " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY CodigoE "
  'SelectDBList DLCurso, AdoAux, sSQL, "Cursos"
  SQLMsg1 = "AÑO LECTIVO: " & Codigo2
  MensajeEncabData = "C A L I F I C A C I O N E S    D E    D I S C I P L I N A"
  If OpcPQBim1.Value Then SQLMsg2 = UCase(OpcPQBim1.Caption) & " DEL PRIMER QUIMESTRE"
  If OpcPQBim2.Value Then SQLMsg2 = UCase(OpcPQBim2.Caption) & " DEL PRIMER QUIMESTRE"
  If OpcPQ.Value Then SQLMsg2 = UCase(OpcPQ.Caption)
  If OpcSQBim1.Value Then SQLMsg2 = "" & UCase(OpcSQBim1.Caption) & " DEL SEGUNDO QUIMESTRE"
  If OpcSQBim2.Value Then SQLMsg2 = "" & UCase(OpcPQBim2.Caption) & " DEL SEGUNDO QUIMESTRE"
  If OpcSQ.Value Then SQLMsg2 = UCase(OpcPQ.Caption)
  SQLMsg3 = "CURSO: " & Codigo1
  
  Imprimir_Disciplina AdoResumen, VectCodMat, True
  
  sSQL = "SELECT (CodigoE & ' - ' & Detalle) As Cursos,Seccion,Detalle As Paralelo,CodigoE " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY CodigoE "
  SelectDBList DLCurso, AdoAux, sSQL, "Cursos"
  DLCurso.Text = CodigoInv
  DGResumen.Visible = True
End Sub

Private Sub Command2_Click()
  Unload FDisciplinaQuimestre
End Sub

Private Sub Command5_Click()
  Listar_Disciplina_Quimestres
End Sub

'''Private Sub Command2_Click()
'''  RatonReloj
'''  'RutaOrigen = "D:\Mis Archivos Walter\Fotos\BMP\BEBES.BMP"
'''  FVerGrafico.Show
'''End Sub

Private Sub DGResumen_Click()

End Sub

Private Sub Form_Activate()
  'DGResumen.Height = FDisciplinaQuimestre.Height - 1700
  'DGResumen.Width = FDisciplinaQuimestre.Width - 300
  sSQL = "SELECT (CodigoE & ' - ' & Detalle) As Cursos,Seccion,Detalle As Paralelo,CodigoE " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND CodigoE >= '2' " _
       & "ORDER BY CodigoE "
  SelectDBList DLCurso, AdoAux, sSQL, "Cursos"
  RatonNormal
  DLCurso.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoResumen
  ConectarAdodc AdoDisciplina
End Sub

Public Sub Listar_Disciplina_Quimestres()
'Procesamos las Disciplinas de Curso
  For I = 1 To 100
      SumaHoriz(I) = 0
      VectNota(100) = 0
      VectMate(100) = "."
      VectCodMat(100) = "."
  Next I
  RatonReloj
  Contador = 0
  ContNotas = 0
  TipoDoc = Ninguno
  Codigo = SinEspaciosIzq(DLCurso.Text)
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       If Codigo = "" Then Codigo = "."
      .MoveFirst
      .Find ("CodigoE = '" & Codigo & "' ")
       If Not .EOF Then
          TipoDoc = .Fields("Seccion")
          Codigo1 = .Fields("Paralelo")
       End If
   End If
  End With
  sSQL = "DELETE * " _
       & "FROM Balances_Mes " _
       & "WHERE TC = 'DI' " _
       & "AND Codigo = '" & Codigo & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT C.Cliente As Alumno,CE.Detalle As Materia,C.Grupo,CE.C,CE.P,TN.* " _
       & "FROM Clientes As C,Catalogo_Estudiantil As CE,Trans_Notas As TN " _
       & "WHERE C.Grupo = '" & Codigo & "' " _
       & "AND CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.CodMat = TN.CodMat " _
       & "AND TN.Codigo = C.Codigo " _
       & "AND C.Grupo = Mid(CE.CodigoE,1,7) " _
       & "AND CE.Item = TN.Item " _
       & "AND CE.Periodo = TN.Periodo " _
       & "ORDER BY C.Cliente,CE.CodigoE "
  'MsgBox sSQL
  SelectDataGrid DGResumen, AdoResumen, sSQL, , True
  DGResumen.Visible = False
  With AdoResumen.Recordset
   If .RecordCount > 0 Then
       CodigoCliente = .Fields("Codigo")
       Contador = 0
       ContHoriz = 0
       ContVert = 0
       Do While Not .EOF
          If CodigoCliente <> .Fields("Codigo") Then
             ContVert = ContVert + 1
             
             SetAdoAddNew "Balances_Mes"
             SetAdoFields "CodigoC", CodigoCliente
             SetAdoFields "Codigo", Codigo
             SetAdoFields "TC", "DI"
             Sumatoria = 0
             ContHoriz = 0
             Insertar_Disciplina_Curso Contador
             SumaHoriz(ContVert) = Sumatoria
             If ContHoriz = 0 Then ContHoriz = 1
             Entrada = Round(Sumatoria / ContHoriz, 2)
             SetAdoFields "TOTAL", Entrada
             SetAdoUpdate
             'MsgBox CodigoCliente
             For I = 1 To Contador
                 VectNota(I) = 0
             Next I
             CodigoCliente = .Fields("Codigo")
             Contador = 0
          End If
          Contador = Contador + 1
          
          VectMate(Contador) = .Fields("Materia")
          VectCodMat(Contador) = .Fields("CodMat")
          
          
          If OpcPQBim1.Value Then
             If .Fields("CodMat") = "999" Or .Fields("CodMat") = "998" Then
                 VectNota(Contador) = Round(.Fields("PQBim1"))
             Else
                 VectNota(Contador) = Round(.Fields("ConductaPQ1"))
             End If
          End If
          
          If OpcPQBim2.Value Then VectNota(Contador) = .Fields("PQBim2")
          If OpcPQ.Value Then VectNota(Contador) = .Fields("PromPQ")
          
          If OpcSQBim1.Value Then VectNota(Contador) = .Fields("SQBim1")
          If OpcSQBim2.Value Then VectNota(Contador) = .Fields("SQBim2")
          If OpcSQ.Value Then VectNota(Contador) = .Fields("PromSQ")
          
          
          'MsgBox OpcPQBim1.Value & " .."
          'If .Fields("C") Then VectMate(Contador) = VectMate(Contador) & "_"
          'If .Fields("P") Then VectMate(Contador) = VectMate(Contador) & "|"
         .MoveNext
       Loop
       ContVert = ContVert + 1
       SetAdoAddNew "Balances_Mes"
       SetAdoFields "CodigoC", CodigoCliente
       SetAdoFields "Codigo", Codigo
       SetAdoFields "TC", "DI"
       Insertar_Disciplina_Curso Contador
       SumaHoriz(ContVert) = Sumatoria
       If ContHoriz = 0 Then ContHoriz = 1
       Entrada = Round(Sumatoria / ContHoriz, 2)
       SetAdoFields "TOTAL", Entrada
       SetAdoUpdate
       .MoveFirst
   End If
  End With
  Total = 0
  For I = 1 To ContVert
    Total = Total + SumaHoriz(I)
  Next I
  If ContNotas = 0 Then ContNotas = 1
  Total = Round(Total / ContNotas)
  
  sSQL = "SELECT C.Cliente As Alumno "
  For I = 1 To Contador + 3
      sSQL = sSQL & ",D_" & Format(I, "00") & " As [" & VectMate(I) & "] "
  Next I
  sSQL = sSQL & ",TOTAL As PROMEDIO " _
       & "FROM Clientes As C, Balances_Mes As BDM " _
       & "WHERE TC = 'DI' " _
       & "AND BDM.Codigo = '" & Codigo & "' " _
       & "AND BDM.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = BDM.CodigoC " _
       & "ORDER BY C.Sexo DESC,C.Cliente "
  'MsgBox sSQL
  SelectDataGrid DGResumen, AdoResumen, sSQL, , True
  If OpcPQ.Value Or OpcSQ.Value Then
  End If
  'MsgBox Total
  DGResumen.Visible = True
  RatonNormal
End Sub


Public Sub Insertar_Disciplina_Curso(MaxContador As Long)
  sSQL = "SELECT * " _
       & "FROM Trans_Asistencia " _
       & "WHERE Codigo = '" & CodigoCliente & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoDisciplina, sSQL
  Sumatoria = 0
  ContHoriz = 0
  For I = 1 To MaxContador
      SetAdoFields "D_" & Format(I, "00"), VectNota(I)
      If VectNota(I) <> 0 Then
         ContHoriz = ContHoriz + 1
         ContNotas = ContNotas + 1
         Sumatoria = Sumatoria + VectNota(I)
      End If
  Next I
  I = MaxContador
  With AdoDisciplina.Recordset
   If .RecordCount > 0 Then
'''    Contador = Contador + 1
'''    SetAdoFields "D_" & Format(Contador, "00"), .Fields("ConductaPQ1")
'''    VectMate(Contador) = .Fields("ConductaPQ1")
       I = I + 1
       SetAdoFields "D_" & Format(I, "00"), .Fields("PQBFJ1")
       VectMate(I) = "Falta Justificada"
       I = I + 1
       SetAdoFields "D_" & Format(I, "00"), .Fields("PQBFI1")
       VectMate(I) = "Falta Injustificada"
       I = I + 1
       SetAdoFields "D_" & Format(I, "00"), .Fields("PQBA1")
       VectMate(I) = "Atrasos"
   End If
  End With
End Sub
