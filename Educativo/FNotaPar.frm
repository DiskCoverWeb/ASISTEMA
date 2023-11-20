VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FNotasParciales 
   Caption         =   "CATALOGO ESTUDIANTIL"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   12405
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "Grabar Notas Grado"
      Enabled         =   0   'False
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
      Left            =   9345
      TabIndex        =   10
      Top             =   105
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FNotaPar.frx":0000
      Height          =   2640
      Left            =   105
      TabIndex        =   1
      Top             =   4200
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4657
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin MSAdodcLib.Adodc AdoAutorizar 
      Height          =   330
      Left            =   7350
      Top             =   6825
      Visible         =   0   'False
      Width           =   4110
      _ExtentX        =   7250
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
      Caption         =   "Autorizar"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar &Actas de Grado"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   648
      Left            =   10395
      TabIndex        =   9
      Top             =   108
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar &Notas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   648
      Left            =   8400
      TabIndex        =   8
      Top             =   105
      Width           =   855
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
      Height          =   645
      Left            =   11445
      TabIndex        =   6
      Top             =   105
      Width           =   855
   End
   Begin MSDataListLib.DataList DLMaterias 
      Bindings        =   "FNotaPar.frx":0019
      DataSource      =   "AdoMaterias"
      Height          =   2595
      Left            =   7035
      TabIndex        =   4
      Top             =   1155
      Visible         =   0   'False
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   4577
      _Version        =   393216
      ForeColor       =   128
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
   Begin ComctlLib.TreeView TVNivel 
      Height          =   3690
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "<Ctrl+L>: Listar Nomina para las notas, <Ctrl+A>: Ingresar Actas de Grado"
      Top             =   105
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6509
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImgList"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FNotaPar.frx":0033
   End
   Begin MSAdodcLib.Adodc AdoNivel 
      Height          =   330
      Left            =   210
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Nivel"
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
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   105
      Top             =   6825
      Width           =   7260
      _ExtentX        =   12806
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoMaterias 
      Height          =   330
      Left            =   210
      Top             =   1155
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1470
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc AdoActas 
      Height          =   330
      Left            =   210
      Top             =   1785
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Actas"
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
   Begin MSAdodcLib.Adodc AdoCatalogoGrado 
      Height          =   330
      Left            =   210
      Top             =   2100
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "CatalogoGrado"
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
   Begin MSAdodcLib.Adodc AdoCursos 
      Height          =   330
      Left            =   210
      Top             =   2415
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Cursos"
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
   Begin VB.Label LblMaterias 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   7035
      TabIndex        =   7
      Top             =   840
      Width           =   4440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Ctrl+F5> Modificar|<Ctrl+F6> No Modifica|<Ctrl+Ins> Insertar|<Ctrl+B> Buscar|<Ctrl+Supr> Eliminar|<Ctrl+V> Cambio de Valores"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   3885
      Width           =   11355
   End
   Begin VB.Label LblValor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7035
      TabIndex        =   2
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   7035
      TabIndex        =   3
      Top             =   105
      Width           =   1275
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   105
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaPar.frx":034D
            Key             =   "C"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaPar.frx":0577
            Key             =   "N"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaPar.frx":0891
            Key             =   "P"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FNotaPar.frx":0BAB
            Key             =   "M"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FNotasParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ListarActasAlunmos(Optional OpcList As Boolean)
  If OpcList Then
     sSQL = "SELECT * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "ORDER BY Id_No "
  Else
     sSQL = "SELECT C.Codigo,C.Cliente As Alumno,TN.* " _
          & "FROM Clientes As C,Trans_Actas As TN " _
          & "WHERE TN.CodE = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "ORDER BY C.Sexo DESC,C.Cliente "
  End If
End Sub

Public Sub LlenarCodigos()
Dim nodX As Node
Dim Vector() As Nodo_Arbol
Dim IndVect As Integer
Dim IndVect1 As Integer
  Si_No = False
' Establece propiedades del control ImageList.
  TVNivel.Visible = False
' Crea un árbol con varios objetos Node sin ordenar.
' Establece propiedades del control ImageList.
  TVNivel.Nodes.Clear
  TVNivel.LineStyle = tvwTreeLines
' Crea un árbol con varios objetos Node sin ordenar.
  IndVect = 1
  sSQL = "SELECT CE.*,CM.Materia,CM.C,CM.I,CM.P,C.Cliente As Dirigente " _
       & "FROM Catalogo_Estudiantil As CE, Catalogo_Materias AS CM, Clientes As C " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.Profesor = C.Codigo " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "ORDER BY CE.CodigoE "
  SelectData AdoNivel, sSQL
  With ImgList
   If AdoNivel.Recordset.RecordCount > 0 Then
      ReDim Vector(1 To AdoNivel.Recordset.RecordCount) As Nodo_Arbol
      Do While Not AdoNivel.Recordset.EOF
         Codigo = "C" & AdoNivel.Recordset.Fields("CodigoE")
         CodigoL = AdoNivel.Recordset.Fields("CodigoE")
         TipoDoc = AdoNivel.Recordset.Fields("TC")
         TipoProc = AdoNivel.Recordset.Fields("CodMat")
         Cadena = AdoNivel.Recordset.Fields("Materia")
         Select Case TipoDoc
           Case "M": Cadena = AdoNivel.Recordset.Fields("Materia")
           Case Else
                If AdoCursos.Recordset.RecordCount > 0 Then
                   AdoCursos.Recordset.MoveFirst
                   AdoCursos.Recordset.Find ("Curso = '" & CodigoL & "'")
                   If Not AdoCursos.Recordset.EOF Then Cadena = AdoCursos.Recordset.Fields("Descripcion")
                End If
         End Select
         If Len(Codigo) = 2 Then
            Set nodX = TVNivel.Nodes.Add(, , Codigo, Cadena, .ListImages(1).key, .ListImages(1).key)
         Else
            Select Case TipoDoc
              Case "N": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(2).key, .ListImages(2).key)
              Case "P": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(3).key, .ListImages(3).key)
              Case "M": Set nodX = TVNivel.Nodes.Add(CambioCodigoCtaSup(Codigo), tvwChild, Codigo, Cadena, .ListImages(4).key, .ListImages(3).key)
            End Select
         End If
         Vector(IndVect).Item_Nodo = Codigo
         Vector(IndVect).Codigo_Aux = AdoNivel.Recordset.Fields("Profesor")
         Vector(IndVect).Eliminar = True
         If Vector(IndVect).Codigo_Aux = CodigoUsuario Then
            Codigo4 = Vector(IndVect).Item_Nodo
            Do While Len(Codigo4) > 1
               For IndVect1 = 1 To IndVect
                If Codigo4 = Vector(IndVect1).Item_Nodo Then Vector(IndVect1).Eliminar = False
               Next IndVect1
               Codigo4 = CodigoCuentaSup(Codigo4)
            Loop
         End If
         IndVect = IndVect + 1
         AdoNivel.Recordset.MoveNext
      Loop
      'nodX.EnsureVisible
      TVNivel.Visible = True
      Si_No = True
      RatonNormal
      Command2.SetFocus
   Else
      RatonNormal
      Unload FNotasParciales
   End If
  End With
  If Si_No Then
     For I = TVNivel.Nodes.Count To 1 Step -1
         Set nodX = TVNivel.Nodes.Item(I)
         If Vector(I).Eliminar Then TVNivel.Nodes.Remove I
     Next I
     TVNivel.Visible = True
     RatonNormal
  Else
     RatonNormal
     MsgBox "No tiene Materias Asignadas"
     End
  End If
End Sub

Public Sub EliminarCta()
  Codigo1 = Mid(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  Cadena = SinEspaciosIzq(TVNivel.SelectedItem.key)
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo like '" & Codigo1 & "' ")
       If Not .EOF Then
          Mensajes = "Esta seguro que desea eliminar la " & vbCrLf _
                   & "Cuenta No. [" & TVNivel.SelectedItem & "]"
          Titulo = "Pregunta de Eliminacion"
          If BoxMensaje = vbYes Then
            .Delete
             TVNivel.Nodes.Remove TVNivel.SelectedItem.Index
          End If
       End If
   End If
  End With
End Sub

Private Sub Command1_Click()
  Mensajes = "Esta seguro de Grabar Notas "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     Actualizar_Notas_del_Curso TipoDoc, CodigoCuentaSup(Codigo)
     Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo)
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     Command1.Enabled = False
     MsgBox "Proceso Terminado"
   End If
End Sub

Private Sub Command2_Click()
  Unload FNotasParciales
End Sub

Private Sub Command3_Click()
  Mensajes = "Esta seguro de Grabar Actas de Grado "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     If SQL_Server Then
        sSQL = "UPDATE Trans_Actas " _
             & "SET Notas = AN.Notas,Evaluacion = AN.Evaluacion," _
             & "Trabajo = AN.Trabajo,Investigacion = AN.Investigacion," _
             & "PromFinal = ROUND((AN.Notas+AN.Evaluacion+AN.Trabajo+AN.Investigacion)/4,2) " _
             & "FROM Trans_Actas as TN,Asiento_A As AN "
     Else
        sSQL = "UPDATE Trans_Actas as TN,Asiento_A As AN " _
             & "SET TN.Notas = AN.Notas,TN.Evaluacion = AN.Evaluacion," _
             & "TN.Trabajo = AN.Trabajo,TN.Investigacion = AN.Investigacion," _
             & "TN.PromFinal = ROUND((AN.Notas+AN.Evaluacion+AN.Trabajo+AN.Investigacion)/4,2) "
     End If
     sSQL = sSQL & "WHERE AN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND AN.CodigoU = '" & CodigoUsuario & "' " _
          & "AND TN.Codigo = AN.Codigo "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     ListarActasAlunmos True
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     Command3.Enabled = False
     MsgBox "Proceso Terminado"
   End If
End Sub


Private Sub DGDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyReturn Then
     AdoDetalle.Recordset.MoveNext
     If AdoDetalle.Recordset.EOF Then AdoDetalle.Recordset.MoveFirst
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     'Codigo
     Codigo1 = DGDetalle.Columns(0).Text
     Codigo2 = DGDetalle.Columns(1).Text
     Mensajes = "Eliminar al Alumno(a): " & UCase(Codigo2) & vbCrLf & vbCrLf _
              & "Codigo del Alumno(a): " & Codigo1 & vbCrLf & vbCrLf _
              & "Curso del Plantel: " & Codigo
     Titulo = "PREGUNTA DE ELIMINACION DE ALUMNOS"
     If BoxMensaje = vbYes Then
        RatonReloj
        sSQL = "DELETE * " _
             & "FROM Trans_Notas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & Codigo1 & "' " _
             & "AND CodE = '" & Codigo & "' "
        ConectarAdoExecute sSQL
        sSQL = "DELETE * " _
             & "FROM Trans_Actas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Codigo = '" & Codigo1 & "' " _
             & "AND CodE = '" & Codigo & "' "
        ConectarAdoExecute sSQL
        sSQL = "UPDATE Clientes " _
             & "SET Grupo = 'RET' " _
             & "WHERE Codigo = '" & Codigo1 & "' "
        ConectarAdoExecute sSQL
        RatonNormal
     End If
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     SQLMsg2 = ""
     Cadena = Mid(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
     Cadena = CambioCodigoCtaSup(Cadena)
     With AdoNivel.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("CodigoE like '" & Cadena & "' ")
          If Not .EOF Then
             MensajeEncabData = .Fields("Detalle")
          End If
      End If
     End With
     Cuadricula = True
     SQLMsg3 = "PROFESOR(A): " & UCase(NombreUsuario)
     SQLMsg2 = "MATERIA: " & TVNivel.SelectedItem
     If Si_No And Evaluar = True Then SQLMsg2 = SQLMsg2 & ": ESCRITOS DE GRADO"
     Imprimir_Nomina_Notas AdoDetalle, AdoAutorizar, , Command10.Enabled
  End If
End Sub

Private Sub DLMaterias_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nodX As Node
  If KeyCode = vbKeyEscape Then DLMaterias.Visible = False
  If KeyCode = vbKeyReturn Then
     Codigo = SinEspaciosDer(DLMaterias.Text)
     Cuenta = Mid(DLMaterias.Text, 1, Len(DLMaterias.Text) - Len(Codigo) - 2)
     Codigo1 = Mid(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
     sSQL = "SELECT * " _
          & "FROM Catalogo_Estudiantil " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'M' " _
          & "AND Mid(CodigoE,1," & Len(Codigo1) & ") = '" & Codigo1 & "' " _
          & "ORDER BY CodigoE "
     'MsgBox sSQL
     SelectAdodc AdoAux, sSQL
     Si_No = True
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             If .Fields("CodMat") = Codigo Then Si_No = False
             Codigo2 = Mid(.Fields("CodigoE"), Len(.Fields("CodigoE")) - 1, 2)
             Codigo2 = Format(Val(Codigo2) + 1, "00")
             'MsgBox "-> " & Codigo2
            .MoveNext
          Loop
      Else
          Codigo2 = "01"
      End If
     End With
     'MsgBox Codigo2
     If Si_No Then
        Set nodX = TVNivel.Nodes.Add("C" & Codigo1, tvwChild, "C" & Codigo1 & "." & Codigo2, Cuenta, ImgList.ListImages(4).key, ImgList.ListImages(4).key)
        SetAdoAddNew "Catalogo_Estudiantil"
        SetAdoFields "TC", "M"
        SetAdoFields "CodMat", Codigo
        SetAdoFields "CodigoE", Codigo1 & "." & Codigo2
        SetAdoFields "Detalle", Cuenta
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        TVNivel.Refresh
       'LlenarCodigos
        sSQL = "SELECT * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY CodigoE "
        SelectData AdoNivel, sSQL
     End If
     LblMaterias.Caption = ""
     DLMaterias.Visible = False
     TVNivel.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Curso "
  SelectAdodc AdoCursos, sSQL
  LlenarCodigos
  DGDetalle.Height = MDI_Y_Max - DGDetalle.Top - 300
  DGDetalle.width = MDI_X_Max - DGDetalle.Left
  Label1.width = MDI_X_Max - DGDetalle.Left
  AdoDetalle.Top = DGDetalle.Top + DGDetalle.Height
  AdoAutorizar.Top = DGDetalle.Top + DGDetalle.Height
  sSQL = "SELECT * " _
       & "FROM Catalogo_Periodo_Lectivo " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoAutorizar, sSQL
  With AdoAutorizar.Recordset
   If .RecordCount > 0 Then
       Si_No = .Fields("NPQP1") Or .Fields("NPQP2") Or .Fields("NPQEX")
       Si_No = Si_No Or .Fields("NSQP1") Or .Fields("NSQP2") Or .Fields("NSQEX")
       Si_No = Si_No Or .Fields("NTQP1") Or .Fields("NTQP2") Or .Fields("NTQEX")
       Si_No = Si_No Or .Fields("NSUPL") Or .Fields("NGRADO")
   End If
  End With
  RatonNormal
  If Si_No = False Then
     MsgBox "ADVERTENCIA: " & vbCrLf & vbCrLf _
           & "                          NO SE PUEDE INGRESAR NOTAS " & vbCrLf & vbCrLf _
           & "                          SOLICITE AUTORIZACION A SECRETARIA              " & vbCrLf
     Unload FNotasParciales
  End If
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoNivel
  ConectarAdodc AdoActas
  ConectarAdodc AdoCursos
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoMaterias
  ConectarAdodc AdoAutorizar
  ConectarAdodc AdoCatalogoGrado
End Sub

Private Sub TVNivel_DblClick()
  SiguienteControl
End Sub

Private Sub TVNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  Command1.Enabled = False
  Command10.Enabled = False
  Cadena = Mid(TVNivel.SelectedItem.key, 2, Len(TVNivel.SelectedItem.key) - 1)
  With AdoNivel.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("CodigoE like '" & Cadena & "' ")
       If Not .EOF Then
          Codigo = .Fields("CodigoE")
          TipoDoc = .Fields("CodMat")
          TipoCta = .Fields("TC")
          Evaluar = .Fields("NG")
          Cuenta = .Fields("Materia")
          Codigo4 = CodigoCuentaSup(Codigo)
       End If
   End If
  End With
  
  If CtrlDown And KeyCode = vbKeyInsert And TipoCta = "P" Then
     LblMaterias.Caption = TVNivel.SelectedItem
     DLMaterias.Visible = True
     DLMaterias.SetFocus
  End If
  
  If CtrlDown And KeyCode = vbKeyL And TipoCta = "P" Then
     sSQL = "SELECT TN.Codigo,Cliente As Alumno,Direccion As Curso " _
          & "FROM Clientes As C,Trans_Notas As TN " _
          & "WHERE TN.CodE = '" & Codigo & "' " _
          & "AND TN.Item = '" & NumEmpresa & "' " _
          & "AND TN.Periodo = '" & Periodo_Contable & "' " _
          & "AND C.Codigo = TN.Codigo " _
          & "GROUP BY TN.Codigo,Cliente,Direccion,Sexo " _
          & "ORDER BY Sexo DESC,Cliente "
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem
  End If
  
 'Listamos las notas de los alumnos
  If CtrlDown And KeyCode = vbKeyL And TipoCta = "M" Then
     Leer_Notas_Parciales TipoDoc, CodigoCuentaSup(Codigo)
     Command1.Enabled = True
     Command3.Enabled = False
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem & " DEL " & CodigoCuentaSup(Codigo)
     Listar_Notas_Alunmos AdoAutorizar, TipoDoc, CodigoCuentaSup(Codigo)
    'MsgBox Si_No & vbCrLf & Evaluar
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     If Si_No And Evaluar = False Then
        MsgBox "No Esta Asignado"
        Unload FNotasParciales
     End If
     Command1.Enabled = True
  End If
  
  If CtrlDown And KeyCode = vbKeyA And TipoCta = "P" Then
     'MsgBox Codigo
     If Mid(Codigo, 1, 4) = "3.03" Then
        Command3.Enabled = True
        Command1.Enabled = False
     sSQL = "DELETE * " _
          & "FROM Asiento_A " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     ListarActasAlunmos
     SelectAdodc AdoDetalle, sSQL
     DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem
     'If AdoDetalle.Recordset.RecordCount <= 0 Then
        sSQL = "SELECT Codigo,Cliente As Alumno,Direccion As Curso " _
             & "FROM Clientes " _
             & "WHERE Grupo = '" & Codigo & "' " _
             & "ORDER BY Cliente "
        SelectData AdoAux, sSQL
        If AdoAux.Recordset.RecordCount <> AdoDetalle.Recordset.RecordCount Then
           'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoDetalle.Recordset.RecordCount
           Do While Not AdoAux.Recordset.EOF
              Si_No = False
              Codigo1 = AdoAux.Recordset.Fields("Codigo")
              If AdoDetalle.Recordset.RecordCount > 0 Then
                 AdoDetalle.Recordset.MoveFirst
                 AdoDetalle.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
                 If Not AdoDetalle.Recordset.EOF Then Si_No = True
              End If
              If Si_No = False Then
                 'MsgBox Mid(Codigo, Len(Codigo) - 1, 2)
                 SetAdoAddNew "Trans_Actas"
                 SetAdoFields "Id_No", Val(Mid(Codigo, Len(Codigo) - 1, 2))
                 SetAdoFields "Codigo", Codigo1
                 SetAdoFields "Periodo", Periodo_Contable
                 SetAdoFields "Item", NumEmpresa
                 SetAdoUpdate
              End If
              AdoAux.Recordset.MoveNext
           Loop
        End If
     'End If
     ListarActasAlunmos
     SelectAdodc AdoAux, sSQL
     Contador = 1
     If AdoAux.Recordset.RecordCount > 0 Then
        Do While Not AdoAux.Recordset.EOF
           SetAdoAddNew "Asiento_A"
           SetAdoFields "Id_No", CByte(Contador)
           SetAdoFields "Codigo", AdoAux.Recordset.Fields("Codigo")
           SetAdoFields "Alumno", AdoAux.Recordset.Fields("Alumno")
           SetAdoFields "Notas", AdoAux.Recordset.Fields("Notas")
           SetAdoFields "Trabajo", AdoAux.Recordset.Fields("Trabajo")
           SetAdoFields "Investigacion", AdoAux.Recordset.Fields("Investigacion")
           SetAdoFields "Evaluacion", AdoAux.Recordset.Fields("Evaluacion")
           SetAdoFields "Item", NumEmpresa
           SetAdoFields "CodigoU", CodigoUsuario
           SetAdoUpdate
           Contador = Contador + 1
           AdoAux.Recordset.MoveNext
        Loop
     End If
     ListarActasAlunmos True
     SelectDataGrid DGDetalle, AdoDetalle, sSQL
     Else
        MsgBox "Este paralelo no es valido"
     End If
  End If
  
 'Insertar las notas de Grado del los alumnos
  If CtrlDown And KeyCode = vbKeyG And TipoCta = "M" Then
     sSQL = "DELETE * " _
          & "FROM Asiento_NG " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodMat = '" & TipoDoc & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     sSQL = "SELECT CodMat,Detalle,CodigoE,Item,Periodo " _
          & "FROM Catalogo_Examen_Grado " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = 'M' " _
          & "AND Mid(CodigoE,1," & Len(Codigo4) & ") = '" & Codigo4 & "' " _
          & "AND CodMat = '" & TipoDoc & "' " _
          & "ORDER BY CodigoE "
     SelectAdodc AdoCatalogoGrado, sSQL
     If AdoCatalogoGrado.Recordset.RecordCount > 0 Then
        sSQL = "SELECT C.Codigo,C.Cliente " _
             & "FROM Clientes As C,Clientes_Matriculas As CM " _
             & "WHERE CM.Grupo_No = '" & Codigo4 & "' " _
             & "AND C.Codigo = CM.Codigo " _
             & "AND CM.Item = '" & NumEmpresa & "' " _
             & "AND CM.Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY C.Cliente "
        SelectData AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Contador = 1
             Do While Not .EOF
                Real1 = 0
                sSQL = "SELECT * " _
                     & "FROM Trans_Notas_Grado " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND Periodo = '" & Periodo_Contable & "' " _
                     & "AND Codigo = '" & .Fields("Codigo") & "' " _
                     & "AND Mid(CodE,1," & Len(Codigo4) & ") = '" & Codigo4 & "' " _
                     & "AND CodMat = '" & TipoDoc & "' "
                SelectAdodc AdoAux, sSQL
                If AdoAux.Recordset.RecordCount > 0 Then
                   Real1 = AdoAux.Recordset.Fields("Examen")
                End If
                SetAdoAddNew "Asiento_NG"
                SetAdoFields "Id_No", Contador
                SetAdoFields "Codigo", .Fields("Codigo")
                SetAdoFields "Alumno", .Fields("Cliente")
                SetAdoFields "CodMat", TipoDoc
                SetAdoFields "CodE", Codigo4
                SetAdoFields "Examen", Real1
                SetAdoFields "Item", NumEmpresa
                SetAdoUpdate
                Contador = Contador + 1
               .MoveNext
             Loop
         End If
        End With
        sSQL = "SELECT Id_No,Alumno,Examen,CodMat,Item,Codigo,CodE,CodigoU " _
             & "FROM Asiento_NG " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodMat = '" & TipoDoc & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' " _
             & "ORDER BY Id_No "
        SQLDec = ""
        SelectDataGrid DGDetalle, AdoDetalle, sSQL
        DGDetalle.Caption = "ALUMNOS DEL: " & TVNivel.SelectedItem & " DEL " & CodigoCuentaSup(Codigo)
        Command10.Enabled = True
     Else
        MsgBox "Esta Materia no esta asignada para examen de Grado"
     End If
  End If
  'PresionoEnter KeyCode
End Sub

Private Sub Command10_Click()
    sSQL = "SELECT * " _
         & "FROM Asiento_NG " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodMat = '" & TipoDoc & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY Id_No "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Real1 = .Fields("Examen")
            CodigoCli = .Fields("Codigo")
            Codigo1 = .Fields("CodE")
            Codigo2 = .Fields("CodMat")
            sSQL = "SELECT * " _
                 & "FROM Trans_Notas_Grado " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo = '" & CodigoCli & "' " _
                 & "AND CodE = '" & Codigo1 & "' " _
                 & "AND CodMat = '" & Codigo2 & "' "
            SelectAdodc AdoActas, sSQL
            If AdoActas.Recordset.RecordCount > 0 Then
               AdoActas.Recordset.Fields("Examen") = Real1
               AdoActas.Recordset.Update
            Else
               SetAdoAddNew "Trans_Notas_Grado"
               SetAdoFields "Id_No", Contador
               SetAdoFields "Codigo", CodigoCli
               SetAdoFields "CodMat", Codigo2
               SetAdoFields "CodE", Codigo1
               SetAdoFields "Examen", Real1
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
            End If
           .MoveNext
         Loop
     End If
    End With
    MsgBox "Proceso de Grabación terminado"
End Sub


