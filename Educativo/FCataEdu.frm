VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FCataEdu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CATALOGO ESTUDIANTIL"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Catalogo de &Materias"
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
      Left            =   9555
      Picture         =   "FCataEdu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1995
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Todos los &Paralelos"
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
      Left            =   9555
      Picture         =   "FCataEdu.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1050
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pensun &Educativo"
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
      Left            =   9555
      Picture         =   "FCataEdu.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   9555
      Picture         =   "FCataEdu.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3885
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Catalogo"
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
      Left            =   9555
      Picture         =   "FCataEdu.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2940
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGCatalogo 
      Bindings        =   "FCataEdu.frx":1AB2
      Height          =   6525
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   11509
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
   Begin MSAdodcLib.Adodc AdoCatalogo 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   9360
      _ExtentX        =   16510
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
      Caption         =   "Adodc1"
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
      Left            =   1995
      Top             =   315
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   210
      Top             =   315
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
      Caption         =   "Aux1"
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
End
Attribute VB_Name = "FCataEdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PointColor
Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
Dim AnchoDeLinea As Single
Dim IJ As Integer
Dim AnchoMaximo As Single
Dim AltoMaximo As Single
 
'''Public Sub Grabar_Disciplinas()
'''  RatonReloj
'''  PictTexto = PictTexto _
'''            & "FALTAS JUSTIF " & vbCrLf _
'''            & "FALTAS INJUSTIF " & vbCrLf _
'''            & "ATRASOS " & vbCrLf _
'''            & "PROMEDIO DISCIPLINA " & vbCrLf
'''  Codigo = Mid$(Codigo4, 1, 1) & Mid$(Codigo4, 3, 2) & Mid$(Codigo4, 6, 2)
'''  Pagina = 1: Msg = ""
'''  PCol = 0.1: PFil = 0
'''  PictLetrasH.ScaleMode = vbCentimeters
'''  PictLetrasV.ScaleMode = vbCentimeters
'''  PictLetrasH.width = AnchoMaximo: PictLetrasH.Height = AltoMaximo
'''  PictLetrasV.Height = AnchoMaximo: PictLetrasV.width = AltoMaximo
'''  PictLetrasH.AutoRedraw = True: PictLetrasH.Picture = LoadPicture()
'''  PictLetrasV.AutoRedraw = True: PictLetrasV.Picture = LoadPicture()
'''  PictLetrasH.FontBold = False: PictLetrasV.FontBold = False
'''  PictLetrasH.Font = TipoLetra: PictLetrasV.Font = TipoLetra
'''  PictLetrasH.FontSize = PorteLetra: PictLetrasV.FontSize = PorteLetra
'''  AltoLetra = PictLetrasH.TextHeight("H") + 0.3
'''  PictTexto = Mid$(PictTexto, 1, Len(PictTexto) - 2)
'''  For I = 1 To Len(PictTexto)
'''    If Mid$(PictTexto, I, 1) = vbCr Then
'''       PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.2
'''       PictLetrasH.Print Msg
'''       Pagina = Pagina + 1
'''       Msg = ""
'''       PFil = PFil + AltoLetra
'''       I = I + 2
'''    End If
'''    Msg = Msg & Mid$(PictTexto, I, 1)
'''  Next I
'''  Texto = SinEspaciosIzq(Msg)
'''  PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.2
'''  PictLetrasH.Print Msg
'''  PFil = 0.7
'''  For I = 1 To Pagina
'''      PictLetrasH.Line (0, PFil)-(AnchoMaximo, PFil), QBColor(0)
'''      PFil = PFil + AltoLetra
'''  Next I
'''  JR = 0
'''  Do While JR <= PictLetrasH.ScaleHeight
'''     IR = 0
'''     Do While IR <= PictLetrasH.ScaleWidth
'''        PointColor = PictLetrasH.Point(IR, JR)
'''        If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''        IR = IR + (PorteLetra / 720)
'''     Loop
'''     JR = JR + (PorteLetra / 720)
'''  Loop
'''  PictLetrasV.Line (0, 0)-(Round(PictLetrasV.width, 2), Round(PictLetrasV.Height, 2)), QBColor(0), B
'''  PictLetrasV.Line (0.01, 0.01)-(Round(PictLetrasV.width - 0.01, 2), Round(PictLetrasV.Height - 0.01, 2)), QBColor(0), B
'''  SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\DISCIPLINA\D" & Codigo & ".BMP"
'''  RatonNormal
'''End Sub

Private Sub Command1_Click()
  SQLMsg1 = Anio_Lectivo
  ImprimirAdodc AdoCatalogo, 1, 8
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
  
  sSQL = "SELECT CE.CodMat,CE.CodigoE,CC.Descripcion As Materia,0 As C,0 As P,0 As I,0 As NG,CE.CodMatP,CE.Item,CE.Periodo " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Cursos As CC " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.Item = CC.Item " _
       & "AND CE.Periodo = CC.Periodo " _
       & "AND CE.CodigoE = CC.Curso " _
       & "AND Mid$(CE.CodigoE,9,2) <> '99' " _
       & "UNION " _
       & "SELECT CE.CodMat,CE.CodigoE,CM.Materia,CM.C,CM.P,CM.I,CE.NG,CE.CodMatP,CE.Item,CE.Periodo " _
       & "FROM Catalogo_Estudiantil As CE,Catalogo_Materias As CM " _
       & "WHERE CE.Item = '" & NumEmpresa & "' " _
       & "AND CE.Periodo = '" & Periodo_Contable & "' " _
       & "AND CE.CodMat <> '.' " _
       & "AND CE.Item = CM.Item " _
       & "AND CE.Periodo = CM.Periodo " _
       & "AND CE.CodMat = CM.CodMat " _
       & "AND Mid$(CE.CodigoE,9,2) <> '99' " _
       & "ORDER BY CE.CodigoE "

  SelectDataGrid DGCatalogo, AdoCatalogo, sSQL
  MensajeEncabData = "C A T A L O G O     D E L     P L A N T E L"
End Sub

'Disciplinas
'''Private Sub Command5_Click()
'''Contador = 0
'''sSQL = "SELECT * " _
'''     & "FROM Catalogo_Estudiantil " _
'''     & "WHERE Item = '" & NumEmpresa & "' " _
'''     & "AND Periodo = '" & Periodo_Contable & "' " _
'''     & "AND TC = 'M' " _
'''     & "ORDER BY CodigoE  "
'''SelectAdodc AdoAux, sSQL
'''RatonReloj
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''PictLetrasH.AutoRedraw = True
'''PictLetrasV.AutoRedraw = True
'''DGCatalogo.Visible = False
'''TipoLetra = TipoArialNarrow
'''PorteLetra = 7.5
'''AltoMaximo = 15.5
'''AnchoMaximo = 3
'''Pagina = 0
'''With AdoAux.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''     PictTexto = ""
'''     Codigo4 = Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
'''     Codigo = Mid$(Codigo4, 1, 1) & Mid$(Codigo4, 3, 2) & Mid$(Codigo4, 6, 2)
'''     Do While Not .EOF
'''        NomCta = .Fields("Detalle")
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & Codigo4
'''       'Grababos los paralelos
'''        If Codigo4 <> Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3) Then
'''           Grabar_Disciplinas
'''           FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
'''           Codigo4 = Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
'''           PictTexto = ""
'''        End If
'''        Contador = Contador + 1
'''        I = 1
'''        Do While PictLetrasH.TextWidth(Mid$(.Fields("Detalle"), 1, I)) < (AnchoMaximo - 0.1) _
'''                 And I <= Len(.Fields("Detalle"))
'''           I = I + 1
'''        Loop
'''        PictTexto = PictTexto & Mid$(.Fields("Detalle"), 1, I) & vbCrLf
'''        RatonNormal
'''       .MoveNext
'''     Loop
'''     Grabar_Disciplinas
'''     FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
''' End If
'''End With
'''DGCatalogo.Visible = True
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''RatonNormal
'''MsgBox "Proceso de Paralelos Terminado"
'''End Sub

Private Sub Command4_Click()
  sSQL = "SELECT Curso,Descripcion,Paralelo,Bachiller,Especialidad,Item,Periodo " _
       & "FROM Catalogo_Cursos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Curso "
  SelectDataGrid DGCatalogo, AdoCatalogo, sSQL
  MensajeEncabData = "C A T A L O G O     D E     P A R A L E L O S"
End Sub

Private Sub Command5_Click()
  sSQL = "SELECT * " _
       & "FROM Catalogo_Materias " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Materia "
  SelectDataGrid DGCatalogo, AdoCatalogo, sSQL
  MensajeEncabData = "C A T A L O G O     D E     M A T E R I A S"
End Sub

'''Private Sub Command6_Click()
'''Dim PointColor
'''Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
'''Dim AnchoDeLinea As Single
'''Dim IJ As Integer
'''' Empezamos la escritura de las letras
'''  RatonReloj
'''  PictLetrasH.Visible = False
'''  PictLetrasV.Visible = False
'''  DGCatalogo.Visible = False
'''
''' 'Archivo para Primaria
'''  PictLetrasH.AutoRedraw = True
'''  PictLetrasV.AutoRedraw = True
'''  PictLetrasH.Picture = LoadPicture()
'''  PictLetrasV.Picture = LoadPicture()
'''  PictLetrasH.FontBold = True
'''  PictLetrasV.FontBold = True
'''  PictLetrasH.Font = TipoArial
'''  PictLetrasV.Font = TipoArial
'''  PictLetrasH.FontSize = 8
'''  PictLetrasV.FontSize = 8
''' 'Texto Vertical
'''  AltoLetra = PictLetrasH.TextHeight("H")
'''  AnchoMax = 2.5
'''  PictLetrasH.Width = AnchoMax
'''  PictLetrasH.Height = 17
'''  PictLetrasV.Height = 2.5
'''  PictLetrasV.Width = 17
'''  PictLetrasH.FontBold = False
'''  PCol = 0.1: PFil = 7.2
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 1 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 2 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "SUMA       "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  'PictLetrasH.Print "           "
'''  PictLetrasH.Print "EXAMEN     "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 1 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 2 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "SUMA       "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  'PictLetrasH.Print "           "
'''  PictLetrasH.Print "EXAMEN     "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.35
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "FINAL      "
'''  PFil = PFil + 0.4
''' 'Transformamos las letras de horizontal a vertical
'''  IR = 0
'''  Do While IR < PictLetrasH.ScaleWidth
'''     JR = 0
'''     Do While JR < PictLetrasH.ScaleHeight
'''        PointColor = PictLetrasH.Point(IR, JR)
'''        If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''        JR = JR + 0.006
'''     Loop
'''     IR = IR + 0.006
'''  Loop
''' 'Todas las Letras y lineas Normales
'''  PCol = 7
'''  For I = 1 To 10
'''      Select Case I
'''        Case 1, 6: PictLetrasV.Line (PCol, 0.01)-(PCol, 2.5), QBColor(0)
'''        Case Else: PictLetrasV.Line (PCol, 0.6)-(PCol, 2.5), QBColor(0)
'''      End Select
'''      PCol = PCol + 0.8
'''  Next I
'''  PictLetrasV.Line (PCol, 0.01)-(PCol, 2.5), QBColor(0)
'''  PictLetrasV.Line (7, 0.6)-(PCol, 0.6), QBColor(0)
'''  PictLetrasV.Line (0.01, 0.7)-(7, 0.7), QBColor(0)
'''  PictLetrasV.Line (0.01, 1.4)-(7, 1.4), QBColor(0)
'''
'''  AnchoDeLinea = PictLetrasH.Width + 0.1
'''  PosLinea = 0.01
'''  PictLetrasV.FontBold = True
'''  PictLetrasV.FontSize = 20
'''  PictLetrasV.CurrentX = 0.8
'''  PictLetrasV.CurrentY = 1.5
'''  PictLetrasV.Print "M A T E R I A S"
'''
'''  PictLetrasV.FontSize = 10
'''  PictLetrasV.CurrentX = 7.4
'''  PictLetrasV.CurrentY = 0.1
'''  PictLetrasV.Print "PRIMER QUIMESTRE"
'''
'''  PictLetrasV.CurrentX = 11.15
'''  PictLetrasV.CurrentY = 0.1
'''  PictLetrasV.Print "SEGUNDO QUIMESTRE"
'''
'''  PictLetrasV.Line (0, 0)-(17, 2.5), QBColor(0), B
'''  PictLetrasV.Line (0.009, 0.009)-(16.9, 2.45), QBColor(0), B
'''  SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\PRIMARIA.BMP"
'''
''' 'Archivo para Secundaria
'''  PictLetrasH.AutoRedraw = True
'''  PictLetrasV.AutoRedraw = True
'''  PictLetrasH.Picture = LoadPicture()
'''  PictLetrasV.Picture = LoadPicture()
'''  PictLetrasH.FontBold = True
'''  PictLetrasV.FontBold = True
'''  PictLetrasH.Font = TipoArial
'''  PictLetrasV.Font = TipoArial
'''  PictLetrasH.FontSize = 8
'''  PictLetrasV.FontSize = 8
''' 'Texto Vertical
'''  AltoLetra = PictLetrasH.TextHeight("H")
'''  AnchoMax = 2.5
'''  PictLetrasH.Width = AnchoMax
'''  PictLetrasH.Height = 17
'''  PictLetrasV.Height = 2.5
'''  PictLetrasV.Width = 17
'''  PCol = 0.1: PFil = 7.2
'''  PictLetrasH.FontBold = False
'''
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 1 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 2 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "SUMA       "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "           "
'''  'PictLetrasH.Print "EXAMEN     "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 1 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "BIMESTRE 2 "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "SUMA       "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "           "
'''  'PictLetrasH.Print "EXAMEN     "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "SUPLETORIO "
'''  PFil = PFil + 0.8
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "PROMEDIO   "
'''  PFil = PFil + 0.35
'''  PictLetrasH.CurrentX = PCol
'''  PictLetrasH.CurrentY = PFil
'''  PictLetrasH.Print "FINAL      "
'''  PFil = PFil + 0.4
''' 'Transformamos las letras de horizontal a vertical
'''  IR = 0
'''  Do While IR < PictLetrasH.ScaleWidth
'''     JR = 0
'''     Do While JR < PictLetrasH.ScaleHeight
'''        PointColor = PictLetrasH.Point(IR, JR)
'''        If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''        JR = JR + 0.006
'''     Loop
'''     IR = IR + 0.006
'''  Loop
''' 'Todas las Letras y lineas Normales
'''  PCol = 7
'''  For I = 1 To 11
'''      Select Case I
'''        Case 1, 6, 11: PictLetrasV.Line (PCol, 0.01)-(PCol, 2.5), QBColor(0)
'''        Case Else: PictLetrasV.Line (PCol, 0.6)-(PCol, 2.5), QBColor(0)
'''      End Select
'''      PCol = PCol + 0.8
'''  Next I
'''  PictLetrasV.Line (PCol, 0.01)-(PCol, 2.5), QBColor(0)
'''  PictLetrasV.Line (7, 0.6)-(PCol - 0.8, 0.6), QBColor(0)
'''  PictLetrasV.Line (0.01, 0.7)-(7, 0.7), QBColor(0)
'''  PictLetrasV.Line (0.01, 1.4)-(7, 1.4), QBColor(0)
'''
'''  AnchoDeLinea = PictLetrasH.Width + 0.1
'''  PosLinea = 0.01
'''  PictLetrasV.FontBold = True
'''  PictLetrasV.FontSize = 20
'''  PictLetrasV.CurrentX = 0.8
'''  PictLetrasV.CurrentY = 1.5
'''  PictLetrasV.Print "M A T E R I A S"
'''
'''  PictLetrasV.FontSize = 10
'''  PictLetrasV.CurrentX = 7.4
'''  PictLetrasV.CurrentY = 0.1
'''  PictLetrasV.Print "PRIMER QUIMESTRE"
'''
'''  PictLetrasV.CurrentX = 11.15
'''  PictLetrasV.CurrentY = 0.1
'''  PictLetrasV.Print "SEGUNDO QUIMESTRE"
'''
'''  PictLetrasV.Line (0, 0)-(17, 2.5), QBColor(0), B
'''  PictLetrasV.Line (0.009, 0.009)-(16.9, 2.45), QBColor(0), B
'''  SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\SECUNDARIA.BMP"
'''  RatonNormal
'''  PictLetrasH.Visible = False
'''  PictLetrasV.Visible = False
'''  DGCatalogo.Visible = True
'''End Sub

Private Sub DGCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyM Then
     MsgBox "PUEDE MODIFICAR LOS DATOS" & vbCrLf _
            & "1= SI/YES" & vbCrLf _
            & "0 = NO/NOT" & vbCrLf
     DGCatalogo.AllowUpdate = True
  End If
  If CtrlDown And KeyCode = vbKeyC Then CodigoC = Trim(Mid$(DGCatalogo.Columns(1).Text, 1, 7))
  If CtrlDown And KeyCode = vbKeyV Then
     CodigoB = UCase(Trim(InputBox("VA HA COPIAR DEL CURSO: " & CodigoC & vbCrLf & vbCrLf & "AL CURSO:", "COPIAR MATERIAS DE UN CURSO A OTRO", CodigoC)))
     If (Len(CodigoB) = 7) And (CodigoB <> CodigoC) Then
        sSQL = "DELETE * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Mid$(CodigoE,1,7) = '" & CodigoB & "' "
        ConectarAdoExecute sSQL
        
        sSQL = "DELETE * " _
             & "FROM Catalogo_Cursos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Curso = '" & CodigoB & "' "
        ConectarAdoExecute sSQL
        
        sSQL = "SELECT * " _
             & "FROM Catalogo_Cursos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Curso = '" & CodigoC & "' "
        SelectAdodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             SetAdoAddNew "Catalogo_Cursos"
             SetAdoFields "Curso", CodigoB
             SetAdoFields "Descripcion", .Fields("Descripcion")
             SetAdoFields "Paralelo", .Fields("Paralelo")
             SetAdoFields "Bachiller", .Fields("Bachiller")
             SetAdoFields "Especialidad", .Fields("Especialidad")
             SetAdoFields "Curso_Superior", .Fields("Curso_Superior")
             SetAdoFields "Ciclo", .Fields("Ciclo")
             SetAdoFields "Seccion", .Fields("Seccion")
             SetAdoUpdate
         End If
        End With
        
        sSQL = "SELECT * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Mid$(CodigoE,1,7) = '" & CodigoC & "' " _
             & "ORDER BY CodigoE  "
        SelectAdodc AdoAux, sSQL
        With AdoAux.Recordset
         If .RecordCount > 0 Then
             Trans_No = 1
             Do While Not .EOF
                SetAdoAddNew "Catalogo_Estudiantil"
                SetAdoFields "TC", .Fields("TC")
                SetAdoFields "Orden", .Fields("Orden")
                SetAdoFields "CodMat", .Fields("CodMat")
                If Len(.Fields("CodigoE")) = 7 Then
                    SetAdoFields "CodigoE", CodigoB
                Else
                    SetAdoFields "CodigoE", CodigoB & "." & Format(Trans_No, "00")
                    Trans_No = Trans_No + 1
                End If
                SetAdoFields "CodMatP", .Fields("CodMatP")
                SetAdoFields "Id_No", Trans_No
                SetAdoUpdate
               .MoveNext
             Loop
             MsgBox "Proceso Terminado"
         End If
        End With
     End If
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     CodigoB = Trim(Mid$(DGCatalogo.Columns(0).Text, 1, 7))
     Titulo = "Pregunta de Eliminación"
     Mensajes = "Esta seguro Eliminar el Curso: " & CodigoB
     If BoxMensaje = vbYes Then
        sSQL = "DELETE * " _
             & "FROM Catalogo_Estudiantil " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Mid$(CodigoE,1,7) = '" & CodigoB & "' "
        ConectarAdoExecute sSQL
        
        sSQL = "DELETE * " _
             & "FROM Catalogo_Cursos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Curso = '" & CodigoB & "' "
        ConectarAdoExecute sSQL
     End If
  End If
  If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto FCataEdu, AdoCatalogo
  If CtrlDown And KeyCode = vbKeyA Then Actualizar_Malla_Cursos
End Sub

Private Sub Form_Activate()
'''  If NombreUsuario = "Administrador de Red" Then
'''     Command4.Enabled = True
'''     Command5.Enabled = True
'''     Command6.Enabled = True
'''     Command18.Enabled = True
'''  Else
'''     Command4.Enabled = False
'''     Command5.Enabled = False
'''     Command6.Enabled = False
'''     Command18.Enabled = False
'''  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCataEdu
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  ConectarAdodc AdoCatalogo
End Sub

'''Private Sub Command18_Click()
'''Dim PointColor
'''Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
'''Dim AnchoDeLinea As Single
'''Dim IJ As Integer
'''' Empezamos la escritura de las letras
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''PictLetrasH.AutoRedraw = True
'''PictLetrasV.AutoRedraw = True
'''DGCatalogo.Visible = False
'''Contador = 0
'''sSQL = "SELECT * " _
'''     & "FROM Catalogo_Materias " _
'''     & "WHERE Item = '" & NumEmpresa & "' " _
'''     & "AND Periodo = '" & Periodo_Contable & "' " _
'''     & "ORDER BY CodMat "
'''SelectAdodc AdoAux, sSQL
'''RatonReloj
'''With AdoAux.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''     Do While Not .EOF
'''        Codigo = .Fields("CodMat")
'''        NomCta = .Fields("Materia")
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & NomCta & ":"
'''
'''        PictLetrasH.AutoRedraw = True
'''        PictLetrasH.Picture = LoadPicture()
'''        PictLetrasV.AutoRedraw = True
'''        PictLetrasV.Picture = LoadPicture()
'''        PictLetrasH.FontBold = True: PictLetrasV.FontBold = True
'''        PictLetrasH.Font = TipoComicSans: PictLetrasV.Font = TipoComicSans
'''        PictLetrasH.FontSize = 10: PictLetrasV.FontSize = 10
'''
'''        PictTexto = "Quimestres" & vbCrLf _
'''                  & "Promedio Global" & vbCrLf _
'''                  & "Examen Supletorio" & vbCrLf _
'''                  & "Promedio Total" & vbCrLf
'''        AltoLetra = PictLetrasH.TextHeight(PictTexto)
'''        AnchoMax = PictLetrasH.TextWidth(PictTexto)
'''        PictLetrasH.Width = Round(AnchoMax) + 1.45
'''        PictLetrasH.Height = Round(AltoLetra) + 0.85
'''        PictLetrasV.Height = PictLetrasH.Width
'''        PictLetrasV.Width = PictLetrasH.Height
'''
'''        PictLetrasH.FontBold = True: PictLetrasV.FontBold = True
'''        PictLetrasH.Font = TipoComicSans: PictLetrasV.Font = TipoComicSans
'''        PictLetrasH.FontSize = 10: PictLetrasV.FontSize = 10
'''
'''        AltoLetra = PictLetrasH.TextHeight(Mid$(PictTexto, 1, 1))
'''        PictLetrasH.Line (AnchoMax + 0.1, 0)-(AnchoMax + 0.1, PictLetrasH.Height - 0.09), QBColor(0)
'''        PCol = 0.1: Msg = "": PFil = 0.1
'''        PictLetrasH.FontBold = False
'''        For I = 1 To Len(PictTexto)
'''         If Mid$(PictTexto, I, 1) = vbCr Then
'''            PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil
'''            PictLetrasH.Print Msg
'''            PictLetrasH.Line (0, PFil + 0.8)-(AnchoMax + 0.1, PFil + 0.8), QBColor(0)
'''           'MsgBox Msg
'''            Msg = ""
'''            PFil = PFil + 0.45 + AltoLetra
'''            I = I + 2
'''         End If
'''         Msg = Msg & Mid$(PictTexto, I, 1)
'''        Next I
'''        RatonReloj
'''        IR = 0
'''        Do While IR < PictLetrasH.ScaleWidth
'''           JR = 0
'''           Do While JR < PictLetrasH.ScaleHeight
'''              PointColor = PictLetrasH.Point(IR, JR)
'''              If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''              JR = JR + 0.005
'''           Loop
'''           IR = IR + 0.005
'''        Loop
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & NomCta & ": (" & IR & ")(" & JR & ")(" & PointColor & ")"
'''        AnchoDeLinea = PictLetrasH.Width + 0.1
'''        PosLinea = 0.01
'''        PictLetrasH.FontBold = True
'''        Texto = SinEspaciosIzq(NomCta)
'''        PictLetrasV.CurrentX = 0.1
'''        PictLetrasV.CurrentY = 0.01
'''        PictLetrasV.Print Texto
'''
'''        PictLetrasV.CurrentX = 0.1
'''        PictLetrasV.CurrentY = 0.45
'''        PictLetrasV.Print Mid$(NomCta, Len(Texto) + 2, Len(NomCta))
'''        PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.09, PictLetrasV.Height - 0.05), QBColor(0), B
'''
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & RutaSistema & "\FORMATOS\MATERIAS\M" & Codigo & ".BMP"
'''        SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\MATERIAS\M" & Codigo & ".BMP"
'''        RatonNormal
'''        'Beep
'''        'MsgBox "Ok"
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & NomCta & ":"
'''        Contador = Contador + 1
'''       .MoveNext
'''     Loop
''' End If
'''End With
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''DGCatalogo.Visible = True
'''RatonNormal
'''MsgBox "Proceso de Materias Terminado"
'''End Sub

'''Private Sub Command4_Click()
'''Dim PointColor
'''Dim PictTexto, Texto, CampoTexto, LineaDeTexto As String
'''Dim AnchoDeLinea As Single
'''Dim IJ As Integer
'''Dim AnchoMaximo As Single
'''Dim AltoMaximo As Single
'''
'''
'''' Empezamos la escritura de las letras
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''PictLetrasH.AutoRedraw = True
'''PictLetrasV.AutoRedraw = True
'''DGCatalogo.Visible = False
'''Contador = 0
'''sSQL = "SELECT * " _
'''     & "FROM Catalogo_Estudiantil " _
'''     & "WHERE Item = '" & NumEmpresa & "' " _
'''     & "AND Periodo = '" & Periodo_Contable & "' " _
'''     & "AND TC = 'M' " _
'''     & "ORDER BY CodigoE  "
'''SelectAdodc AdoAux, sSQL
'''RatonReloj
'''PictTexto = ""
'''AnchoMaximo = 0
'''AltoMaximo = 0
'''
'''With AdoAux.Recordset
''' If .RecordCount > 0 Then
'''    .MoveFirst
'''     TipoLetra = TipoTimes
'''     PorteLetra = 8
'''     AltoMaximo = 15.5
'''     AnchoMaximo = 3
'''     'MsgBox Codigo & vbCrLf & PictTexto & vbCrLf & AnchoMaximo & vbCrLf & AltoMaximo
'''     PictTexto = ""
'''     Codigo4 = Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
'''     Codigo = Mid$(Codigo4, 1, 1) & Mid$(Codigo4, 3, 2) & Mid$(Codigo4, 6, 2)
'''     Do While Not .EOF
'''        NomCta = .Fields("Detalle")
'''        FCataEdu.Caption = Format(Contador / .RecordCount, "00%")
'''       'Grababos los paralelos
'''        If Codigo4 <> Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3) Then
'''           FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & Codigo4
'''           PictTexto = PictTexto & "PROMEDIO " & vbCrLf
'''           Codigo = Mid$(Codigo4, 1, 1) & Mid$(Codigo4, 3, 2) & Mid$(Codigo4, 6, 2)
'''           PictLetrasH.AutoRedraw = True
'''           PictLetrasH.Picture = LoadPicture()
'''           PictLetrasV.AutoRedraw = True
'''           PictLetrasV.Picture = LoadPicture()
'''           PictLetrasH.FontBold = False: PictLetrasV.FontBold = False
'''           PictLetrasH.Font = TipoLetra: PictLetrasV.Font = TipoLetra
'''           PictLetrasH.FontSize = PorteLetra: PictLetrasV.FontSize = PorteLetra
'''
'''           PictLetrasH.Width = AnchoMaximo
'''           PictLetrasH.Height = AltoMaximo
'''           PictLetrasV.Height = PictLetrasH.Width
'''           PictLetrasV.Width = PictLetrasH.Height
'''           PictTexto = Mid$(PictTexto, 1, Len(PictTexto) - 2)
'''
'''           'MsgBox PictTexto
'''
'''           AltoLetra = PictLetrasH.TextHeight(Mid$(PictTexto, 1, 1))
'''           PCol = 0.1: Msg = "": PFil = 0.01
'''           For I = 1 To Len(PictTexto)
'''             If Mid$(PictTexto, I, 1) = vbCr Then
'''                Texto = SinEspaciosIzq(Msg)
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
'''                PictLetrasH.Print Texto
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
'''                PictLetrasH.Print Mid$(Msg, Len(Texto) + 2, Len(Msg))
'''                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
'''                Msg = ""
'''                PFil = PFil + 0.4 + AltoLetra
'''                I = I + 2
'''
'''             End If
'''             Msg = Msg & Mid$(PictTexto, I, 1)
'''           Next I
'''
'''           Texto = SinEspaciosIzq(Msg)
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
'''                PictLetrasH.Print Texto
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
'''                PictLetrasH.Print Mid$(Msg, Len(Texto) + 2, Len(Msg))
'''                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
'''
'''          'MsgBox Texto
'''
'''          RatonReloj
'''          JR = 0
'''          Do While JR <= PictLetrasH.ScaleHeight
'''             IR = 0
'''             Do While IR <= PictLetrasH.ScaleWidth
'''                PointColor = PictLetrasH.Point(IR, JR)
'''                If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''                IR = IR + (PorteLetra / 720)
'''             Loop
'''             JR = JR + (PorteLetra / 720)
'''          Loop
'''          PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.01, PictLetrasV.Height - 0.01), QBColor(0), B
'''
'''          FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
'''          SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
'''          'MsgBox "."
'''          Codigo4 = Mid$(.Fields("CodigoE"), 1, Len(.Fields("CodigoE")) - 3)
'''          PictTexto = ""
'''        End If
'''        Contador = Contador + 1
'''        PictTexto = PictTexto & .Fields("Detalle") & vbCrLf
'''        RatonNormal
'''       .MoveNext
'''     Loop
''' 'Beep
'''           FCataEdu.Caption = Format(Contador / .RecordCount, "00%") & " - " & Codigo4
'''           PictTexto = PictTexto & "PROMEDIO " & vbCrLf
'''           Codigo = Mid$(Codigo4, 1, 1) & Mid$(Codigo4, 3, 2) & Mid$(Codigo4, 6, 2)
'''           'MsgBox Codigo
'''           PictLetrasH.AutoRedraw = True
'''           PictLetrasH.Picture = LoadPicture()
'''           PictLetrasV.AutoRedraw = True
'''           PictLetrasV.Picture = LoadPicture()
'''           PictLetrasH.FontBold = False: PictLetrasV.FontBold = False
'''           PictLetrasH.Font = TipoLetra: PictLetrasV.Font = TipoLetra
'''           PictLetrasH.FontSize = PorteLetra: PictLetrasV.FontSize = PorteLetra
'''
'''           PictLetrasH.Width = AnchoMaximo: PictLetrasH.Height = AltoMaximo
'''           PictLetrasV.Width = AltoMaximo:  PictLetrasV.Height = AnchoMaximo
'''           AltoLetra = PictLetrasH.TextHeight(Mid$(PictTexto, 1, 1))
'''           PCol = 0.1: Msg = "": PFil = 0.01
'''           For I = 1 To Len(PictTexto)
'''             If Mid$(PictTexto, I, 1) = vbCr Then
'''                Texto = SinEspaciosIzq(Msg)
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
'''                PictLetrasH.Print Texto
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
'''                PictLetrasH.Print Mid$(Msg, Len(Texto) + 2, Len(Msg))
'''                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
'''                Msg = ""
'''                PFil = PFil + 0.4 + AltoLetra
'''                I = I + 2
'''             End If
'''             Msg = Msg & Mid$(PictTexto, I, 1)
'''           Next I
'''          Texto = SinEspaciosIzq(Msg)
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.1
'''                PictLetrasH.Print Texto
'''                PictLetrasH.CurrentX = PCol: PictLetrasH.CurrentY = PFil + 0.4
'''                PictLetrasH.Print Mid$(Msg, Len(Texto) + 2, Len(Msg))
'''                PictLetrasH.Line (0, PFil + 0.8)-(AnchoMaximo, PFil + 0.8), QBColor(0)
'''          RatonReloj
'''          IR = 0
'''          Do While IR <= PictLetrasH.ScaleWidth
'''             JR = 0
'''             Do While JR <= PictLetrasH.ScaleHeight
'''                PointColor = PictLetrasH.Point(IR, JR)
'''                If PointColor <> CLng(&HFFFFFF) Then PictLetrasV.PSet (JR, PictLetrasH.ScaleWidth - IR), PointColor
'''                JR = JR + (PorteLetra / 720)
'''             Loop
'''             IR = IR + (PorteLetra / 720)
'''          Loop
'''          PictLetrasV.Line (0.01, 0.01)-(PictLetrasV.Width - 0.01, PictLetrasV.Height - 0.01), QBColor(0), B
'''          SavePicture PictLetrasV.Image, RutaSistema & "\FORMATOS\PARALELO\P" & Codigo & ".BMP"
''' End If
'''End With
'''DGCatalogo.Visible = True
'''PictLetrasH.Visible = False
'''PictLetrasV.Visible = False
'''RatonNormal
'''MsgBox "Proceso de Paralelos Terminado"
'''End Sub

Public Sub Actualizar_Malla_Cursos()
Dim Orden_N As Byte
 'Actualizar Grados
 If ClaveAdministrador Then
    RatonReloj
    sSQL = "SELECT * " _
         & "FROM Catalogo_Estudiantil " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' "
    SelectAdodc AdoAux1, sSQL
    RatonReloj
    sSQL = "SELECT CodE, Id_No, CodMat, CodMatP, Orden " _
         & "FROM Trans_Notas " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "GROUP BY CodE, Id_No, CodMat, CodMatP, Orden " _
         & "ORDER BY CodE, Id_No, CodMat, CodMatP, Orden "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            RatonReloj
            CodMat = .Fields("CodMat")
            CodMatP = .Fields("CodMatP")
            Codigo = .Fields("CodE")
            Orden_N = .Fields("Orden")
            
            Codigo2 = Codigo & "." & Format(.Fields("Id_No"), "00")
            Codigo1 = Codigo
            
            RatonReloj
            If Len(Codigo) = 7 Then
              'Inserto Catalogo cursos
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Cursos " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Curso = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Cursos"
                  SetAdoFields "Curso", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
              'Insertar Catalogo Estudiantil Paralelos
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Estudiantil " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND CodigoE = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Estudiantil"
                  SetAdoFields "TC", "P"
                  SetAdoFields "CodigoE", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoFields "Orden", Orden_N
                  SetAdoUpdate
               End If
              'Insertar Catalogo Estudiantil Seccion
               Codigo1 = CambioCodigoCtaSup(Codigo1)
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Cursos " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Curso = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Cursos"
                  SetAdoFields "Curso", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
               
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Estudiantil " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND CodigoE = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Estudiantil"
                  SetAdoFields "TC", "N"
                  SetAdoFields "CodigoE", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
              'Insertar Catalogo Estudiantil Nivel
               Codigo1 = CambioCodigoCtaSup(Codigo1)
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Cursos " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND Curso = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Cursos"
                  SetAdoFields "Curso", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
               
               sSQL = "SELECT * " _
                    & "FROM Catalogo_Estudiantil " _
                    & "WHERE Periodo = '" & Periodo_Contable & "' " _
                    & "AND Item = '" & NumEmpresa & "' " _
                    & "AND CodigoE = '" & Codigo1 & "' "
               SelectAdodc AdoAux1, sSQL
               If AdoAux1.Recordset.RecordCount <= 0 Then
                  SetAdoAddNew "Catalogo_Estudiantil"
                  SetAdoFields "TC", "C"
                  SetAdoFields "CodigoE", Codigo1
                  SetAdoFields "Periodo", Periodo_Contable
                  SetAdoFields "Item", NumEmpresa
                  SetAdoUpdate
               End If
               RatonReloj
               If .Fields("Id_No") > 0 Then
                    sSQL = "SELECT * " _
                         & "FROM Catalogo_Estudiantil " _
                         & "WHERE Periodo = '" & Periodo_Contable & "' " _
                         & "AND Item = '" & NumEmpresa & "' " _
                         & "AND CodigoE = '" & Codigo2 & "' "
                    SelectAdodc AdoAux1, sSQL
                    If AdoAux1.Recordset.RecordCount <= 0 Then
                       SetAdoAddNew "Catalogo_Estudiantil"
                       SetAdoFields "TC", "M"
                       SetAdoFields "CodigoE", Codigo2
                       SetAdoFields "CodMat", CodMat
                       SetAdoFields "CodMatP", CodMatP
                       SetAdoFields "Periodo", Periodo_Contable
                       SetAdoFields "Item", NumEmpresa
                       SetAdoFields "Id_No", .Fields("Id_No")
                       SetAdoUpdate
                    End If
                End If
            End If
           .MoveNext
         Loop
     End If
    End With
    
    sSQL = "SELECT * " _
         & "FROM Catalogo_Cursos " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "ORDER BY Curso "
    SelectAdodc AdoAux, sSQL
    RatonReloj
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            sSQL = "SELECT * " _
                 & "FROM Catalogo_Cursos " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Curso = '" & .Fields("Curso") & "' " _
                 & "AND Len(Descripcion) > 1 " _
                 & "ORDER BY Periodo "
            SelectAdodc AdoAux1, sSQL
            If AdoAux1.Recordset.RecordCount > 0 Then
              .Fields("Descripcion") = AdoAux1.Recordset.Fields("Descripcion")
              .Fields("Paralelo") = AdoAux1.Recordset.Fields("Paralelo")
              .Fields("Bachiller") = AdoAux1.Recordset.Fields("Bachiller")
              .Fields("Especialidad") = AdoAux1.Recordset.Fields("Especialidad")
              .Fields("Ciclo") = AdoAux1.Recordset.Fields("Ciclo")
              .Fields("Seccion") = AdoAux1.Recordset.Fields("Seccion")
              .Fields("Titulo") = AdoAux1.Recordset.Fields("Titulo")
              .Fields("Tipo_Titulo") = AdoAux1.Recordset.Fields("Tipo_Titulo")
              .Fields("Codigo_Titulo") = AdoAux1.Recordset.Fields("Codigo_Titulo")
              .Fields("Curso_Superior") = AdoAux1.Recordset.Fields("Curso_Superior")
              .Update
            End If
           .MoveNext
         Loop
     End If
    End With
    sSQL = "SELECT CodMat " _
         & "FROM Catalogo_Materias " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodMat = '" & Ninguno & "' "
    SelectAdodc AdoAux, sSQL
    RatonReloj
    With AdoAux.Recordset
     If .RecordCount <= 0 Then
         SetAdoAddNew "Catalogo_Materias"
         SetAdoFields "CodMat", Ninguno
         SetAdoFields "Materia", Ninguno
         SetAdoFields "Periodo", Periodo_Contable
         SetAdoFields "Item", NumEmpresa
         SetAdoUpdate
     End If
    End With
    RatonReloj
    sSQL = "UPDATE Catalogo_Estudiantil " _
         & "SET Orden = 9 " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND CodMat IN ('997','998','999') "
    ConectarAdoExecute sSQL
    RatonNormal
    MsgBox "Proceso terminado, vuelva general notas en blanco"
 End If
End Sub

