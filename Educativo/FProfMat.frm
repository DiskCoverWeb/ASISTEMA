VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form FProfesorMateria 
   Caption         =   "ASIGNACION DE PROFESORES A MATERIA"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11220
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "&Todas los Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   13440
      Picture         =   "FProfMat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   105
      Width           =   1800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Email Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9765
      Picture         =   "FProfMat.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1590
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Email Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8295
      Picture         =   "FProfMat.frx":0AD8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Todas las Materias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11445
      Picture         =   "FProfMat.frx":0F1A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1905
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4740
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "LISTADO DE PROFESORES Y MATERIAS"
      TabPicture(0)   =   "FProfMat.frx":15B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DGMaterias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ALUMNOS POR MATERIA"
      TabPicture(1)   =   "FProfMat.frx":15CC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGAlumnos"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DGMaterias 
         Bindings        =   "FProfMat.frx":15E8
         Height          =   2745
         Left            =   105
         TabIndex        =   7
         ToolTipText     =   $"FProfMat.frx":1602
         Top             =   420
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   4842
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTADO DE PROFESORES Y MATERIAS"
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
      Begin MSDataGridLib.DataGrid DGAlumnos 
         Bindings        =   "FProfMat.frx":169B
         Height          =   2745
         Left            =   -74895
         TabIndex        =   8
         ToolTipText     =   "<Ctrl+I> Insertar / <Ctrl+E> Desactivar"
         Top             =   420
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   4842
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTADO DE PROFESORES Y MATERIAS"
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
   End
   Begin VB.CommandButton Command3 
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
      Height          =   750
      Left            =   15330
      Picture         =   "FProfMat.frx":16B4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DCProfesores 
      Bindings        =   "FProfMat.frx":1F7E
      DataSource      =   "AdoProfesores"
      Height          =   315
      Left            =   5775
      TabIndex        =   3
      Top             =   945
      Visible         =   0   'False
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   16275
      Picture         =   "FProfMat.frx":1F9A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   1365
      Top             =   1995
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
      Caption         =   "Materias"
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
   Begin MSAdodcLib.Adodc AdoProfesores 
      Height          =   330
      Left            =   1365
      Top             =   2310
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
      Caption         =   "Profesores"
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
      Left            =   1365
      Top             =   2625
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
   Begin MSDataListLib.DataCombo DCGrupo 
      Bindings        =   "FProfMat.frx":2864
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   1365
      Top             =   2940
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
      Caption         =   "Grupo"
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
   Begin MSAdodcLib.Adodc AdoAlumnos 
      Height          =   330
      Left            =   1365
      Top             =   3255
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
      Caption         =   "Alumnos"
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
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Seleccione el Curso a Listar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Visible         =   0   'False
      Width           =   5685
   End
End
Attribute VB_Name = "FProfesorMateria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TodoCurso As Boolean
Dim Mensaje_General

Private Sub Command1_Click()
 Unload FProfesorMateria
End Sub

Private Sub Command2_Click()
  With AdoGrupo.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Email_Materias_Curso SinEspaciosIzq(.Fields("Cursos"))
         .MoveNext
       Loop
   End If
  End With
  MsgBox "PROCESO TERMINADO"
End Sub

Private Sub Command3_Click()
  RatonReloj
  DGMaterias.Visible = False
  MensajeEncabData = "LISTADO DE MATERIAS CON PROFESORES"
  SQLMsg1 = ""
  ImprimirAdodc AdoMaterias, 1, 9
  DGMaterias.Visible = True
  RatonNormal
End Sub

Private Sub Command4_Click()
  Email_Materias_Curso SinEspaciosIzq(DCGrupo)
  MsgBox "PROCESO TERMINADO"
End Sub

Private Sub Command5_Click()
  Email_Informes_Curso SinEspaciosIzq(DCGrupo)
  MsgBox "PROCESO TERMINADO"
End Sub

Private Sub Command6_Click()
  With AdoGrupo.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Email_Informes_Curso SinEspaciosIzq(.Fields("Cursos"))
         .MoveNext
       Loop
   End If
  End With
  MsgBox "PROCESO TERMINADO"
End Sub

Private Sub DCGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCGrupo_LostFocus()
  Listar_Curso SinEspaciosIzq(DCGrupo)
End Sub

Private Sub DCProfesores_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCProfesores_LostFocus()
  CodigoCliente = Ninguno
  CodigoB = DCProfesores.Text
  With AdoProfesores.Recordset
   If .RecordCount > 0 Then
       If CodigoB = "" Then CodigoB = Ninguno
      .MoveFirst
      .Find ("Cliente = '" & CodigoB & "' ")
       If Not .EOF Then CodigoCliente = .Fields("Codigo")
   End If
  End With
  Listar_Materias
  DCProfesores.Visible = False
  Label1.Visible = False
  Listar_Curso SinEspaciosIzq(DCGrupo)
  DGMaterias.SetFocus
End Sub

Private Sub DGMaterias_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  TipoCta = DGMaterias.Columns(0)  'TC
  CodigoA = DGMaterias.Columns(1)  'Curso
  CodigoP = DGMaterias.Columns(2)  'Materia
  CodigoB = DGMaterias.Columns(3)  'Profesor
  Codigos = DGMaterias.Columns(4)  'Codigo de Materia
  Codigo = DGMaterias.Columns(5)   'Email
  CodigoCliente = Ninguno
  CodigoB = Trim(Replace(CodigoB, ".", " "))
  CodigoP = Trim(Replace(CodigoP, ".", " "))
 'MsgBox CodigoA & vbCrLf & CodigoP & vbCrLf & CodigoB & vbCrLf & Codigos
  TodoCurso = False
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGMaterias.Visible = False
     GenerarDataTexto FProfesorMateria, AdoMaterias
     DGMaterias.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyE Then
     DGMaterias.Visible = False
     CodigoCliente = Ninguno
     Listar_Materias
     Listar_Curso CodigoA
     DGMaterias.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyI Then
     DGMaterias.Visible = False
     Label1.Caption = " " & CodigoA & " - " & CodigoP & ":"
     Label1.Visible = True
     DCProfesores.Visible = True
     DCProfesores.SetFocus
     DGMaterias.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyT Then
     DGMaterias.Visible = False
     Label1.Caption = " " & CodigoA & " - " & CodigoP & ":"
     TodoCurso = True
     Label1.Visible = True
     DCProfesores.Visible = True
     DCProfesores.SetFocus
     DGMaterias.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     DGMaterias.Visible = False
     Label1.Caption = " " & CodigoA & " - " & CodigoP & ":"
     MensajeEncabData = "NOMINA DE MATERIAS Y PROFESORES CON DIRIGENTES"
     Cuadricula = True
     ImprimirAdodc AdoMaterias, 1, 9
     Cuadricula = False
     DGMaterias.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyL Then
     CodigoProv = Ninguno
     With AdoProfesores.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente = '" & CodigoB & "' ")
          If Not .EOF Then CodigoProv = .Fields("Codigo")
      End If
     End With
     sSQL = "SELECT C.Cliente,'0' As Nota,C.Grupo,'" & Codigos & "' As CodMat,CM.Codigo,'" & CodigoProv & "' As Codigo_Profesor " _
          & "FROM Clientes As C,Clientes_Matriculas As CM " _
          & "WHERE CM.Item = '" & NumEmpresa & "' " _
          & "AND CM.Periodo = '" & Periodo_Contable & "' " _
          & "AND CM.Grupo_No = '" & CambioCodigoCtaSup(CodigoA) & "' " _
          & "AND CM.Codigo = C.Codigo " _
          & "ORDER BY C.Cliente "
     SelectAdodc AdoAlumnos, sSQL
     Generar_Lista_Materia AdoAlumnos
  End If
  If CtrlDown And KeyCode = vbKeyM Then
     RatonReloj
     CodigoCliente = Ninguno
     If TipoCta = "M" And EsUnEmail(Codigo) And Len(CodigoB) > 1 Then
        Select Case Mid$(CodigoA, 1, 1)
          Case "1": TMail.ListaMail = 1
          Case "2": TMail.ListaMail = 2
          Case "3": TMail.ListaMail = 3
          Case Else: TMail.ListaMail = 0
        End Select
        TMail.Adjunto = Listar_Alumnos_por_Materia(CodigoA, CodigoP, Codigos, CodigoB)
        TMail.Asunto = CodigoP
        TMail.Mensaje = Mensaje_General
        TMail.para = Codigo
        FEnviarCorreos.Show 1
        TMail.ListaMail = 0
     End If
     RatonNormal
  End If
  If CtrlDown And KeyCode = vbKeyR Then
     RatonReloj
     CodigoCliente = Ninguno
     If TipoCta = "M" And EsUnEmail(Codigo) And Len(CodigoB) > 1 Then
        Select Case Mid$(CodigoA, 1, 1)
          Case "1": TMail.ListaMail = 1
          Case "2": TMail.ListaMail = 2
          Case "3": TMail.ListaMail = 3
          Case Else: TMail.ListaMail = 0
        End Select
        TMail.Adjunto = Listar_Informes_Alumnos(CodigoA, CodigoP, Codigos, CodigoB)
        TMail.Asunto = "RECOMENDACIONES Y PLAN DE MEJORA ACADEMICO DE " & CodigoP
        TMail.Mensaje = Mensaje_General
        TMail.para = Codigo
        FEnviarCorreos.Show 1
        TMail.ListaMail = 0
     End If
     RatonNormal
  End If
End Sub

Private Sub Form_Activate()
  SSTab1.Height = MDI_Y_Max - DCProfesores.Top - 400
  SSTab1.width = MDI_X_Max - 100
  
  SSTab1.Tab = 1
  DGAlumnos.Height = SSTab1.Height - DGAlumnos.Top - 100
  DGAlumnos.width = SSTab1.width - DGAlumnos.Left - 100
  
  SSTab1.Tab = 0
  DGMaterias.Height = SSTab1.Height - DGMaterias.Top - 100
  DGMaterias.width = SSTab1.width - DGMaterias.Left - 100
  
  sSQL = "SELECT C.Cliente,CR.* " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND C.Codigo = CR.Codigo " _
       & "ORDER BY C.Cliente "
  SelectDBCombo DCProfesores, AdoProfesores, sSQL, "Cliente"
  
  sSQL = "SELECT (Curso & ' - ' & Descripcion) As Cursos " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Len(Curso)>4 " _
       & "ORDER BY Curso "
  SelectDBCombo DCGrupo, AdoGrupo, sSQL, "Cursos"
  TodoCurso = False
  Mensaje_General = "Estimado Docente, descargue el archivo, guardelo en su computador e" & vbCrLf _
                  & "ingrese las notas de los alumnos y guarde la informacion," & vbCrLf _
                  & "si aparece un mensaje que dice: 'Desea abrir el archivo ahora?' " & vbCrLf _
                  & "Presionar el boton SI, Luego de ingresar las notas envie a este correo el archivo." & vbCrLf & vbCrLf _
                  & "Si presenta un mensaje como este: " & vbCrLf _
                  & "Vista Protegida, Presionar el Boton que dice: ""Habilitar edición"" " & vbCrLf _
                  & " " & vbCrLf _
                  & "Nota: Recuerde que cada Asignatura tiene su propio Codigo,."
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FProfesorMateria
  ConectarAdodc AdoAux
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoAlumnos
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoProfesores
End Sub

Public Sub Listar_Curso(Curso As String)
  sSQL = "SELECT CE.TC,CE.CodigoE as Curso,CM.Materia,C.Cliente As Profesor,CM.CodMat,C.Email,CE.CodMatP " _
       & "FROM Catalogo_Estudiantil As CE, Catalogo_Materias AS CM, Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.TC IN ('M','P') " _
       & "AND Mid$(CE.CodigoE,1,7) = '" & Curso & "' " _
       & "AND CE.Profesor = C.Codigo " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "ORDER BY CE.CodigoE "
  SelectDataGrid DGMaterias, AdoMaterias, sSQL
End Sub

Public Sub Listar_Materias()
  If TodoCurso Then
     sSQL = "UPDATE Catalogo_Estudiantil " _
          & "SET Profesor = '" & CodigoCliente & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC IN ('M','P') " _
          & "AND Mid$(CodigoE,1," & Len(CodigoA) & ") = '" & CodigoA & "' "
  Else
     sSQL = "UPDATE Catalogo_Estudiantil " _
          & "SET Profesor = '" & CodigoCliente & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC IN ('M','P') " _
          & "AND CodigoE = '" & CodigoA & "' "
  End If
  ConectarAdoExecute sSQL
'''  With AdoMaterias.Recordset
'''   If .RecordCount > 0 Then
'''      .MoveFirst
'''      .Find ("CodigoE = '" & CodigoA & "' ")
'''   End If
'''  End With
End Sub

Public Sub Generar_Lista_Materia(AdoList As Adodc)
Dim FileList1 As Long
Dim FileList2 As Long
Dim Traza As String
Dim RutaGeneraFileAlumnos As String
  RatonReloj
 'Abrimo los archivo que vamos ha necesitar
  FileList1 = FreeFile
  Traza = CodigoP & " - 97"
  RutaGeneraFileAlumnos = UCase(RutaSysBases & "\Emails\" & Traza) & ".csv"
  Open RutaGeneraFileAlumnos For Output As #FileList1
  FileList2 = FreeFile
  Traza = CodigoP & " - 2010"
  RutaGeneraFileAlumnos = UCase(RutaSysBases & "\Emails\" & Traza) & ".csv"
  Open RutaGeneraFileAlumnos For Output As #FileList2
  Contador = 0
' Comenzamos a generar el archivo: EMPLEADOS.TXT
  With AdoList.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Print #FileList1, "No.;Alumnos;Nota;Curso;CodMat;Codigo;Codigo_Profesor;"
       Print #FileList2, "No.,Alumnos,Nota,Curso,CodMat,Codigo,Codigo_Profesor,"
       Do While Not .EOF
          Contador = Contador + 1
          Print #FileList1, Format(Contador, "00") & ";";
          Print #FileList1, .Fields("Cliente") & ";";
          Print #FileList1, .Fields("Nota") & ";";
          Print #FileList1, "'" & .Fields("Grupo") & ";";
          Print #FileList1, "'" & .Fields("CodMat") & ";";
          Print #FileList1, "'" & .Fields("Codigo") & ";";
          Print #FileList1, "'" & .Fields("Codigo_Profesor") & ";"
          
          Print #FileList2, Format(Contador, "00") & ",";
          Print #FileList2, .Fields("Cliente") & ",";
          Print #FileList2, .Fields("Nota") & ",";
          Print #FileList2, "'" & .Fields("Grupo") & ",";
          Print #FileList2, "'" & .Fields("CodMat") & ",";
          Print #FileList2, "'" & .Fields("Codigo") & ",";
          Print #FileList2, "'" & .Fields("Codigo_Profesor") & ","
          
         .MoveNext
       Loop
   End If
  End With
  Close #FileList1
  Close #FileList2
  RatonNormal
End Sub

Public Function Listar_Alumnos_por_Materia(Curso As String, _
                                           Materia As String, _
                                           CodigoMateria As String, _
                                           Profesor As String) As String
Dim NFila As Integer
Dim NCelda As Integer
Dim RutaGeneraFile As String
Dim Paralelo As String
Dim NotaExa As Boolean
Dim NotaSup As Boolean
Dim NotaRem As Boolean
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

  RatonReloj
  DGAlumnos.Visible = False
  NotaExa = False
  NotaSup = False
  NotaRem = False
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  
  Profesor = Replace(Profesor, ".", "")
  Profesor = Replace(Profesor, "/", "_")
  Profesor = Replace(Profesor, " ", "_")
  Profesor = UCase(Trim(Profesor))
  
  Materia = Replace(Materia, ".", "")
  Materia = Replace(Materia, "/", "_")
  Materia = Replace(Materia, ":", " ")
  Materia = Replace(Materia, " ", "_")
  Materia = UCase(Trim(Materia))
  
  Paralelo = "N" & Mid$(Curso, 1, 1) & "-" & Trim(Mid$(CambioCodigoCtaSup(Curso), 3, 10))
  Paralelo = Trim(Replace(Paralelo, ".", "_"))
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAux, sSQL
  RatonReloj
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       If .Fields("NPQEX") Then NotaExa = True
       If .Fields("NSQEX") Then NotaExa = True
       If .Fields("NSUPL") Then NotaSup = True
       If .Fields("NREME") Then NotaRem = True
   End If
  End With
  Contador = 0
  RatonReloj
  Codigo1 = Leer_Datos_del_Curso(Curso)
  Codigo1 = Dato_Curso.Especialidad
  RatonReloj
  RutaGeneraFile = RutaSysBases & "\Emails\Notas\" & Paralelo & " " & Materia & " " & Profesor & ".xls"
  If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
  sSQL = "SELECT C.Cliente As Alumno,'0' As TAI,'0' As AIC,'0' As AGC,'0' As L,'0' As Examen," _
       & "TN.CodMat,C.Codigo,C.Sexo,C.Grupo " _
       & "FROM Clientes As C,Clientes_Matriculas As CM,Trans_Notas As TN " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & CambioCodigoCtaSup(Curso) & "' " _
       & "AND TN.CodMat = '" & CodigoMateria & "' " _
       & "AND C.FA <> " & Val(adFalse) & " "
  If NotaSup Or NotaRem Then sSQL = sSQL & "AND TN.PromFinal < " & Nota_Rojo & " "
  sSQL = sSQL _
       & "AND C.Codigo = CM.Codigo " _
       & "AND C.Codigo = TN.Codigo " _
       & "AND CM.Grupo_No = TN.CodE " _
       & "AND CM.Periodo = TN.Periodo " _
       & "AND CM.Item = TN.Item " _
       & "UNION " _
       & "SELECT C.Cliente As Alumno,'0' As TAI,'0' As AIC,'0' As AGC,'0' As L,'0' As Examen," _
       & "TN.CodMat,C.Codigo,C.Sexo,C.Grupo " _
       & "FROM Clientes As C,Clientes_Matriculas As CM,Trans_Notas_Auxiliares As TN " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & CambioCodigoCtaSup(Curso) & "' " _
       & "AND TN.CodMat = '" & CodigoMateria & "' " _
       & "AND C.FA <> " & Val(adFalse) & " "
  If NotaSup Or NotaRem Then sSQL = sSQL & "AND TN.PromFinal < " & Nota_Rojo & " "
  sSQL = sSQL _
       & "AND C.Codigo = CM.Codigo " _
       & "AND C.Codigo = TN.Codigo " _
       & "AND CM.Grupo_No = TN.CodE " _
       & "AND CM.Periodo = TN.Periodo " _
       & "AND CM.Item = TN.Item " _
       & "ORDER BY C.Cliente,C.Sexo "
  SelectDataGrid DGAlumnos, AdoAlumnos, sSQL
  SelectAdodc AdoAlumnos, sSQL
  RatonReloj
  With AdoAlumnos.Recordset
   If .RecordCount > 0 Then
       NFila = 1
       RatonReloj
       If Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
          oSheet.Columns("A").ColumnWidth = 60
          oSheet.Range("A1").value = "E S T U D I A N T E"
          oSheet.Range("N1").value = "PROMEDIO"
          If Mid$(.Fields("Grupo"), 1, 4) <= "1.01" Then
             oSheet.Range("B1").value = "NOTA CUALITATIVA"
             oSheet.Range("C1").value = ""
             oSheet.Range("D1").value = ""
             oSheet.Range("E1").value = ""
             If Val(.Fields("CodMat")) >= 997 Then
                oSheet.Range("F1").value = "DIAS"
                oSheet.Range("G1").value = "FJ"
                oSheet.Range("H1").value = "FI"
                oSheet.Range("I1").value = "ATRASOS"
             Else
                oSheet.Range("F1").value = ""
                oSheet.Range("G1").value = ""
                oSheet.Range("H1").value = ""
                oSheet.Range("I1").value = ""
             End If
          Else
             If Val(.Fields("CodMat")) >= 997 Then
                oSheet.Range("B1").value = "NOTA"
                oSheet.Range("C1").value = ""
                oSheet.Range("D1").value = ""
                oSheet.Range("E1").value = ""
                oSheet.Range("F1").value = "DIAS"
                oSheet.Range("G1").value = "FJ"
                oSheet.Range("H1").value = "FI"
                oSheet.Range("I1").value = "ATRASOS"
             Else
                If NotaExa Then
                   oSheet.Range("B1").value = "Examen Quimestral"
                   oSheet.Range("C1").value = ""
                   oSheet.Range("D1").value = ""
                   oSheet.Range("E1").value = ""
                   oSheet.Range("F1").value = ""
                   oSheet.Range("N1").value = "PROMEDIO"
                ElseIf NotaSup Then
                   oSheet.Range("B1").value = "Supletorio"
                   oSheet.Range("C1").value = ""
                   oSheet.Range("D1").value = ""
                   oSheet.Range("E1").value = ""
                   oSheet.Range("F1").value = ""
                ElseIf NotaRem Then
                   oSheet.Range("B1").value = "Remedial"
                   oSheet.Range("C1").value = ""
                   oSheet.Range("D1").value = ""
                   oSheet.Range("E1").value = ""
                   oSheet.Range("F1").value = ""
                Else
                   oSheet.Range("B1").value = "TAI"
                   oSheet.Range("C1").value = "AIC"
                   oSheet.Range("D1").value = "AGC"
                   oSheet.Range("E1").value = "L"
                   oSheet.Range("F1").value = "Examen"
                   oSheet.Range("G1").value = "FJ"
                   oSheet.Range("H1").value = "FI"
                   oSheet.Range("I1").value = "ATRASOS"
                End If
             End If
          End If
          oSheet.Range("J1").value = "CodMateria"
          oSheet.Range("K1").value = "Codigo"
          oSheet.Range("L1").value = "Sexo"
          oSheet.Range("M1").value = "Curso"
          oSheet.Range("A1:M1").Font.Bold = True
       Else
          oSheet.Range("A1").value = "Alumno"
          oSheet.Range("B1").value = "Nota"
          oSheet.Range("C1").value = "CodMateria"
          oSheet.Range("D1").value = "Codigo"
          oSheet.Range("E1").value = "Sexo"
          oSheet.Range("F1").value = "Curso"
          oSheet.Range("A1:F1").Font.Bold = True
       End If
       NFila = 1
       Do While Not .EOF
          NFila = NFila + 1
          If Mid$(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
             oSheet.Range("A" & CStr(NFila)).value = .Fields("Alumno")
             If NotaExa Then
                oSheet.Range("B" & CStr(NFila)).value = CInt(.Fields("TAI"))
                oSheet.Range("C" & CStr(NFila)).value = ""
                oSheet.Range("D" & CStr(NFila)).value = ""
                oSheet.Range("E" & CStr(NFila)).value = ""
                oSheet.Range("F" & CStr(NFila)).value = ""
                oSheet.Range("N" & CStr(NFila)).value = CInt(.Fields("TAI"))
             Else
                oSheet.Range("B" & CStr(NFila)).value = CInt(.Fields("TAI"))
                oSheet.Range("C" & CStr(NFila)).value = CInt(.Fields("AIC"))
                oSheet.Range("D" & CStr(NFila)).value = CInt(.Fields("AGC"))
                oSheet.Range("E" & CStr(NFila)).value = CInt(.Fields("L"))
                oSheet.Range("F" & CStr(NFila)).value = CInt(.Fields("Examen"))
                oSheet.Range("G" & CStr(NFila)).value = "0"
                oSheet.Range("H" & CStr(NFila)).value = "0"
                oSheet.Range("I" & CStr(NFila)).value = "0"
                oSheet.Cells(NFila, 14).formula = "=SUM(B" & CStr(NFila) & ":" & "F" & CStr(NFila) & ")/5"
             End If
             If Val(.Fields("CodMat")) <= 997 Then
                 oSheet.Range("G" & CStr(NFila)).value = "0"
                 oSheet.Range("H" & CStr(NFila)).value = "0"
                 oSheet.Range("I" & CStr(NFila)).value = "0"
             End If
             oSheet.Range("J" & CStr(NFila)).value = .Fields("CodMat")
             oSheet.Range("K" & CStr(NFila)).value = .Fields("Codigo")
             oSheet.Range("L" & CStr(NFila)).value = .Fields("Sexo")
             oSheet.Range("M" & CStr(NFila)).value = .Fields("Grupo")
          Else
             oSheet.Range("A" & CStr(NFila)).value = .Fields("Alumno")
             oSheet.Range("B" & CStr(NFila)).value = CInt(.Fields("TAI"))
             oSheet.Range("C" & CStr(NFila)).value = .Fields("CodMateria")
             oSheet.Range("D" & CStr(NFila)).value = .Fields("Codigo")
             oSheet.Range("E" & CStr(NFila)).value = .Fields("Sexo")
             oSheet.Range("F" & CStr(NFila)).value = .Fields("Grupo")
          End If
         .MoveNext
       Loop
       
      'Bloqueamos las celdas que no se puden cambiar
       For NCelda = 66 To 76
           oSheet.Columns(Chr(NCelda)).ColumnWidth = 8
       Next NCelda
       For NCelda = 1 To 14
           With oSheet.Cells(1, NCelda)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 41    ' Color fondo = azul '41
               .Font.Size = 9             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NFila + 1, NCelda)  ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 41    ' Color fondo = azul '41
               .Font.Size = 9             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
       Next NCelda
       For NCelda = 2 To NFila
           With oSheet.Cells(NCelda, 1)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 10             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NCelda, 10)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 8             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NCelda, 11)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 8             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NCelda, 12)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 8             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NCelda, 13)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 8             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
       Next NCelda
       oSheet.Unprotect "DiskCoverEducativo"
       oSheet.Range("B2:B" & CStr(NFila)).Locked = False
       oSheet.Range("C2:C" & CStr(NFila)).Locked = False
       oSheet.Range("D2:D" & CStr(NFila)).Locked = False
       oSheet.Range("E2:E" & CStr(NFila)).Locked = False
       oSheet.Range("F2:F" & CStr(NFila)).Locked = False
       oSheet.Range("G2:F" & CStr(NFila)).Locked = False
       oSheet.Range("H2:F" & CStr(NFila)).Locked = False
       oSheet.Range("I2:F" & CStr(NFila)).Locked = False
       oSheet.Protect "DiskCoverEducativo"
      'Save the Workbook and Quit Excel
        
       oBook.SaveAs RutaGeneraFile
       oExcel.Quit
   Else
      RutaGeneraFile = ""
   End If
  End With
  RatonNormal
  DGAlumnos.Visible = True
  Listar_Alumnos_por_Materia = RutaGeneraFile
End Function

Public Function Listar_Informes_Alumnos(Curso As String, _
                                        Materia As String, _
                                        CodigoMateria As String, _
                                        Profesor As String) As String
Dim NFila As Integer
Dim NColumna As Integer
Dim NCelda As Integer
Dim RutaGeneraFile As String
Dim Paralelo As String
Dim NotaExa As Boolean
Dim NotaSup As Boolean
Dim NotaRem As Boolean
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

  RatonReloj
  DGAlumnos.Visible = False
  NotaExa = False
  NotaSup = False
  NotaRem = False
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  
  Profesor = Replace(Profesor, ".", "")
  Profesor = Replace(Profesor, "/", "_")
  Profesor = Replace(Profesor, " ", "_")
  Profesor = UCase(Trim(Profesor))
  
  Materia = Replace(Materia, ".", "")
  Materia = Replace(Materia, "/", "_")
  Materia = Replace(Materia, ":", " ")
  Materia = Replace(Materia, " ", "_")
  Materia = UCase(Trim(Materia))
  
  Paralelo = "RPMA" & Mid$(Curso, 1, 1) & "-" & Trim(Mid$(CambioCodigoCtaSup(Curso), 3, 10))
  Paralelo = Trim(Replace(Paralelo, ".", "_"))
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAux, sSQL
  RatonReloj
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       If .Fields("NPQEX") Then NotaExa = True
       If .Fields("NSQEX") Then NotaExa = True
       If .Fields("NSUPL") Then NotaSup = True
       If .Fields("NREME") Then NotaRem = True
   End If
  End With
  Contador = 0
  RatonReloj
  Codigo1 = Leer_Datos_del_Curso(Curso)
  Codigo1 = Dato_Curso.Especialidad
  RatonReloj
  RutaGeneraFile = RutaSysBases & "\Emails\Notas\" & Paralelo & " " & Materia & " " & Profesor & ".xls"
  If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
  sSQL = "SELECT C.Cliente As Alumno,'0' As TAI,'0' As AIC,'0' As AGC,'0' As L,'0' As Examen," _
       & "TN.CodMat,C.Codigo,C.Sexo,C.Grupo " _
       & "FROM Clientes As C,Clientes_Matriculas As CM,Trans_Notas As TN " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & CambioCodigoCtaSup(Curso) & "' " _
       & "AND TN.CodMat = '" & CodigoMateria & "' "
  If NotaSup Or NotaRem Then sSQL = sSQL & "AND TN.PromFinal < " & Nota_Rojo & " "
  sSQL = sSQL _
       & "AND C.Codigo = CM.Codigo " _
       & "AND C.Codigo = TN.Codigo " _
       & "AND CM.Grupo_No = TN.CodE " _
       & "AND CM.Periodo = TN.Periodo " _
       & "AND CM.Item = TN.Item " _
       & "UNION " _
       & "SELECT C.Cliente As Alumno,'0' As TAI,'0' As AIC,'0' As AGC,'0' As L,'0' As Examen," _
       & "TN.CodMat,C.Codigo,C.Sexo,C.Grupo " _
       & "FROM Clientes As C,Clientes_Matriculas As CM,Trans_Notas_Auxiliares As TN " _
       & "WHERE CM.Item = '" & NumEmpresa & "' " _
       & "AND CM.Periodo = '" & Periodo_Contable & "' " _
       & "AND CM.Grupo_No = '" & CambioCodigoCtaSup(Curso) & "' " _
       & "AND TN.CodMat = '" & CodigoMateria & "' "
  If NotaSup Or NotaRem Then sSQL = sSQL & "AND TN.PromFinal < " & Nota_Rojo & " "
  sSQL = sSQL _
       & "AND C.Codigo = CM.Codigo " _
       & "AND C.Codigo = TN.Codigo " _
       & "AND CM.Grupo_No = TN.CodE " _
       & "AND CM.Periodo = TN.Periodo " _
       & "AND CM.Item = TN.Item " _
       & "ORDER BY C.Cliente,C.Sexo "
  SelectDataGrid DGAlumnos, AdoAlumnos, sSQL
  SelectAdodc AdoAlumnos, sSQL
  RatonReloj
  With AdoAlumnos.Recordset
   If .RecordCount > 0 Then
       NFila = 1
       RatonReloj
       oSheet.Columns("A").ColumnWidth = 60
       oSheet.Range("A1").value = "E S T U D I A N T E"
       oSheet.Range("B1").value = "INFORME AL ALUMNO"
       oSheet.Range("C1").value = "CodMateria"
       oSheet.Range("D1").value = "Codigo"
       oSheet.Range("E1").value = "Sexo"
       oSheet.Range("F1").value = "Curso"
       oSheet.Range("A1:F1").Font.Bold = True
       NFila = 1
       Do While Not .EOF
          NFila = NFila + 1
          oSheet.Range("A" & CStr(NFila)).value = .Fields("Alumno")
          oSheet.Range("B" & CStr(NFila)).value = ""
          oSheet.Range("C" & CStr(NFila)).value = .Fields("CodMat")
          oSheet.Range("D" & CStr(NFila)).value = .Fields("Codigo")
          oSheet.Range("E" & CStr(NFila)).value = .Fields("Sexo")
          oSheet.Range("F" & CStr(NFila)).value = .Fields("Grupo")
         .MoveNext
       Loop
       
      'Bloqueamos las celdas que no se puden cambiar
       oSheet.Columns("A").ColumnWidth = 50
       oSheet.Columns("B").ColumnWidth = 80
       oSheet.Columns("C").ColumnWidth = 9
       oSheet.Columns("D").ColumnWidth = 9
       oSheet.Columns("E").ColumnWidth = 6
       oSheet.Columns("F").ColumnWidth = 6
       
       For NCelda = 1 To 6
           With oSheet.Cells(1, NCelda)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 41    ' Color fondo = azul '41
               .Font.Size = 9             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           With oSheet.Cells(NFila + 1, NCelda)  ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 41    ' Color fondo = azul '41
               .Font.Size = 9             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
       Next NCelda
       For NCelda = 2 To NFila
           With oSheet.Cells(NCelda, 1)    ' seleccionamos la 1ª celda
               .Interior.ColorIndex = 42    ' Color fondo = azul '41
               .Font.Size = 10             ' tamaño de letra
               .Font.Bold = True           ' Fuente en negrita
               .Font.ColorIndex = 2       ' Color fuente = blanco
           End With
           For NColumna = 3 To 6
               With oSheet.Cells(NCelda, NColumna)    ' seleccionamos la 1ª celda
                   .Interior.ColorIndex = 42    ' Color fondo = azul '41
                   .Font.Size = 8             ' tamaño de letra
                   .Font.Bold = True           ' Fuente en negrita
                   .Font.ColorIndex = 2       ' Color fuente = blanco
               End With
           Next NColumna
       Next NCelda
       oSheet.Unprotect "DiskCoverEducativo"
       oSheet.Range("B2:B" & CStr(NFila)).Locked = False
       oSheet.Protect "DiskCoverEducativo"
      'Save the Workbook and Quit Excel
       oBook.SaveAs RutaGeneraFile
       oExcel.Quit
   Else
      RutaGeneraFile = ""
   End If
  End With
  RatonNormal
  DGAlumnos.Visible = True
  Listar_Informes_Alumnos = RutaGeneraFile
End Function

Public Sub Email_Materias_Curso(Curso As String)
Dim EsperarMail As Integer
Dim ContMat As Long
    ContMat = 0
    Listar_Curso Curso
    DGMaterias.Visible = False
    With AdoMaterias.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            CodigoA = .Fields("Curso")   'Curso
            CodigoP = .Fields("Materia")   'Materia
            CodigoB = .Fields("Profesor")   'Profesor
            Codigos = .Fields("CodMat")   'Codigo de Materia
            CodigoB = Trim(Replace(CodigoB, ".", " "))
            CodigoP = Trim(Replace(CodigoP, ".", " "))
            CodigoCliente = Ninguno
            Select Case Mid$(CodigoA, 1, 1)
              Case "1": TMail.ListaMail = 1
              Case "2": TMail.ListaMail = 2
              Case "3": TMail.ListaMail = 3
              Case Else: TMail.ListaMail = 0
            End Select
            If .Fields("TC") = "M" And EsUnEmail(.Fields("Email")) And Len(CodigoB) > 1 Then
                TMail.Adjunto = Listar_Alumnos_por_Materia(CodigoA, CodigoP, Codigos, CodigoB)
                TMail.Asunto = CodigoP
                TMail.Mensaje = Mensaje_General
                TMail.para = .Fields("Email")
                FEnviarCorreos.Show 1
            End If
            ContMat = ContMat + 1
            FProfesorMateria.Caption = Format(ContMat / .RecordCount, "00%") & " " & CodigoA & " - " & CodigoB & " - " & CodigoP
            FProfesorMateria.Refresh
           .MoveNext
         Loop
      End If
    End With
    TMail.ListaMail = 0
    DGMaterias.Visible = True
End Sub

Public Sub Email_Informes_Curso(Curso As String)
Dim EsperarMail As Integer
Dim ContMat As Long
    ContMat = 0
    Listar_Curso Curso
    DGMaterias.Visible = False
    With AdoMaterias.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            CodigoA = .Fields("Curso")   'Curso
            CodigoP = .Fields("Materia")   'Materia
            CodigoB = .Fields("Profesor")   'Profesor
            Codigos = .Fields("CodMat")   'Codigo de Materia
            CodigoB = Trim(Replace(CodigoB, ".", " "))
            CodigoP = Trim(Replace(CodigoP, ".", " "))
            CodigoCliente = Ninguno
            Select Case Mid$(CodigoA, 1, 1)
              Case "1": TMail.ListaMail = 1
              Case "2": TMail.ListaMail = 2
              Case "3": TMail.ListaMail = 3
              Case Else: TMail.ListaMail = 0
            End Select
            If .Fields("TC") = "M" And EsUnEmail(.Fields("Email")) And Len(CodigoB) > 1 Then
                TMail.Adjunto = Listar_Informes_Alumnos(CodigoA, CodigoP, Codigos, CodigoB)
                TMail.Asunto = "RECOMENDACIONES Y PLAN DE MEJORA ACADEMICO DE " & CodigoP
                TMail.Mensaje = Mensaje_General
                TMail.para = .Fields("Email")
                FEnviarCorreos.Show 1
            End If
            ContMat = ContMat + 1
            FProfesorMateria.Caption = Format(ContMat / .RecordCount, "00%") & " " & CodigoA & " - " & CodigoB & " - " & CodigoP
            FProfesorMateria.Refresh
           .MoveNext
         Loop
      End If
    End With
    TMail.ListaMail = 0
    DGMaterias.Visible = True
End Sub

